/*
 * mood.c —— 心情记录本（图形界面版）
 *
 * 功能极简：选一个心情标签 -> 打个 1~10 分 -> 点「记下来」。
 * 数据追加保存在 mood_log.csv（UTF-8 带 BOM，Excel 可直接打开）。
 *
 * UI 参考 ColorOS 7 设计语言：浅色底 + 白色圆角卡片 + 胶囊按钮 + 绿色点缀，
 * 全部 GDI 自绘（无系统默认控件外观），双缓冲防闪烁。
 *
 * 纯 Win32 API，单文件，零第三方依赖。
 *
 * 编译： gcc -O2 -municode -mwindows -Wall -o mood.exe mood.c
 */

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#ifndef _CRT_SECURE_NO_WARNINGS
#define _CRT_SECURE_NO_WARNINGS
#endif

#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <time.h>

/* ---------------- 常量 ---------------- */

#define APP_TITLE   L"心情记录本"
#define DATA_FILE   "mood_log.csv"
#define MAX_RECORD  40000        /* 内存里最多容纳的日期数（够用几十年） */

#define IDC_TAG_BASE    100      /* 心情按钮 100..111 */
#define IDC_SCORE_BASE  200      /* 分数按钮 200..209 */
#define IDC_SAVE        300

static const wchar_t *kTags[12] = {
    L"开心", L"平静", L"兴奋", L"感激", L"放松",
    L"专注", L"疲惫", L"焦虑", L"低落", L"生气",
    L"孤独", L"迷茫"
};
static const int kTagCount = 12;

/* 客户区大小 */
#define CLIENT_W 560
#define CLIENT_H 452

/* ------- 配色（ColorOS 7 风格） ------- */
#define C_BG        RGB(244,246,250)   /* 页面浅灰底 */
#define C_CARD      RGB(255,255,255)   /* 卡片白 */
#define C_CARD_LINE RGB(233,236,241)   /* 卡片描边 */
#define C_ACCENT    RGB(38,194,129)    /* ColorOS 绿 */
#define C_ACCENT_DK RGB(30,170,112)    /* 按下加深 */
#define C_TEXT      RGB(31,36,41)      /* 主文字 */
#define C_TEXT_SUB  RGB(148,155,164)   /* 次要文字 */
#define C_CHIP_BG   RGB(240,243,247)   /* 未选胶囊底 */
#define C_CHIP_DN   RGB(227,232,239)   /* 未选胶囊按下 */
#define C_CHIP_TX   RGB(72,79,88)      /* 未选胶囊文字 */

/* ------- 布局常量 ------- */
#define CARD1_X 24
#define CARD1_Y 94
#define CARD1_W 512
#define CARD1_H 132
#define CARD2_Y 238
#define CARD2_H 84

/* ---------------- 全局状态 ---------------- */

static HWND g_hwnd;
static HWND g_tagbtns[12], g_scorebtns[10], g_save;

static int g_sel_tag   = -1;     /* 0..11 */
static int g_sel_score = 0;      /* 0 表示未选 */

static char g_dates[MAX_RECORD][16];  /* 每条记录的日期 */
static int  g_ndates = 0;

static HFONT g_f_title, g_f_sub, g_f_chip, g_f_chip_b, g_f_btn, g_f_sel, g_f_status;
static HBRUSH g_br_card;    /* 复用的白色刷子（WM_CTLCOLORBTN 用） */

/* ---------------- 日期工具 ---------------- */

static void fmt_day(time_t t, char *buf, size_t sz) {
    struct tm tv;
#ifdef _WIN32
    localtime_s(&tv, &t);
#else
    localtime_r(&t, &tv);
#endif
    strftime(buf, sz, "%Y-%m-%d", &tv);
}

static void today_str(char *buf, size_t sz)       { fmt_day(time(NULL), buf, sz); }
static void day_back(char *buf, size_t sz, int n) { fmt_day(time(NULL) - (time_t)n * 86400, buf, sz); }

static int count_day(const char *d) {
    int c = 0;
    for (int i = 0; i < g_ndates; i++)
        if (strcmp(g_dates[i], d) == 0) c++;
    return c;
}

/* ---------------- 数据读写 ---------------- */

static int file_exists(void) {
    FILE *f = fopen(DATA_FILE, "r");
    if (!f) return 0;
    fclose(f);
    return 1;
}

static void ensure_file(void) {
    if (file_exists()) return;
    FILE *f = fopen(DATA_FILE, "w");
    if (!f) return;
    fputs("\xEF\xBB\xBF" "date,score,mood\n", f);   /* UTF-8 BOM + 表头 */
    fclose(f);
}

static void load_dates(void) {
    g_ndates = 0;
    FILE *f = fopen(DATA_FILE, "r");
    if (!f) return;

    char line[512];
    if (!fgets(line, sizeof line, f)) { fclose(f); return; }   /* 表头（含 BOM），丢弃 */

    while (fgets(line, sizeof line, f) && g_ndates < MAX_RECORD) {
        char *comma = strchr(line, ',');
        if (!comma) continue;
        size_t len = (size_t)(comma - line);
        if (len < 10 || len >= sizeof g_dates[0]) continue;
        memcpy(g_dates[g_ndates], line, len);
        g_dates[g_ndates][len] = '\0';
        g_ndates++;
    }
    fclose(f);
}

/* 追加一条记录；成功返回 1 */
static int append_record(int tag, int score) {
    ensure_file();
    FILE *f = fopen(DATA_FILE, "a");
    if (!f) return 0;

    char date[16], tag8[64];
    today_str(date, sizeof date);
    WideCharToMultiByte(CP_UTF8, 0, kTags[tag], -1, tag8, (int)sizeof tag8, NULL, NULL);

    fprintf(f, "%s,%d,%s\n", date, score, tag8);
    fclose(f);
    return 1;
}

/* 连续记录天数（今天还没记则从昨天起算） */
static int calc_streak(void) {
    char d[16];
    int back = 0;
    today_str(d, sizeof d);
    if (count_day(d) == 0) back = 1;

    int streak = 0;
    while (back <= 3650) {
        day_back(d, sizeof d, back);
        if (count_day(d) == 0) break;
        streak++;
        back++;
    }
    return streak;
}

/* ---------------- 字体 ---------------- */

static HFONT mkfont(int size, int weight) {
    return CreateFontW(size, 0, 0, 0, weight, 0, 0, 0, DEFAULT_CHARSET,
                       0, 0, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
                       L"Microsoft YaHei UI");
}

static void mkfonts(void) {
    g_f_title   = mkfont(-24, FW_BOLD);
    g_f_sub     = mkfont(-14, FW_NORMAL);
    g_f_chip    = mkfont(-15, FW_NORMAL);
    g_f_chip_b  = mkfont(-15, FW_SEMIBOLD);
    g_f_btn     = mkfont(-18, FW_BOLD);
    g_f_sel     = mkfont(-17, FW_SEMIBOLD);
    g_f_status  = mkfont(-13, FW_NORMAL);
}

/* ---------------- 绘制 ---------------- */

/* 居中画一行字 */
static void draw_text_c(HDC hdc, const wchar_t *s, HFONT f, COLORREF color, int y) {
    RECT rc = { 0, y, CLIENT_W, y + 60 };
    HGDIOBJ of = SelectObject(hdc, f);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, color);
    DrawTextW(hdc, s, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);
}

/* 胶囊按钮：圆角 = 高度的一半 */
static void draw_capsule(HDC hdc, RECT rc, COLORREF bg, COLORREF line,
                         const wchar_t *text, HFONT f, COLORREF tcolor) {
    HBRUSH  b   = CreateSolidBrush(bg);
    HPEN    pen = CreatePen(PS_SOLID, 1, line ? line : bg);
    HGDIOBJ ob  = SelectObject(hdc, b);
    HGDIOBJ op  = SelectObject(hdc, pen);
    int h = rc.bottom - rc.top;
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, h, h);

    HGDIOBJ of = SelectObject(hdc, f);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, tcolor);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    SelectObject(hdc, ob);
    SelectObject(hdc, op);
    SelectObject(hdc, of);
    DeleteObject(b);
    DeleteObject(pen);
}

/* 页面底色 + 两张白色圆角卡片 */
static void paint_background(HDC hdc) {
    RECT rc = { 0, 0, CLIENT_W, CLIENT_H };
    HBRUSH bg = CreateSolidBrush(C_BG);
    FillRect(hdc, &rc, bg);
    DeleteObject(bg);

    HBRUSH  card = CreateSolidBrush(C_CARD);
    HPEN    line = CreatePen(PS_SOLID, 1, C_CARD_LINE);
    HGDIOBJ ob   = SelectObject(hdc, card);
    HGDIOBJ op   = SelectObject(hdc, line);
    RoundRect(hdc, CARD1_X, CARD1_Y, CARD1_X + CARD1_W, CARD1_Y + CARD1_H, 20, 20);
    RoundRect(hdc, CARD1_X, CARD2_Y, CARD1_X + CARD1_W, CARD2_Y + CARD2_H, 20, 20);
    SelectObject(hdc, ob);
    SelectObject(hdc, op);
    DeleteObject(card);
    DeleteObject(line);
}

/* 标题 / 已选 / 状态文字 */
static void paint_texts(HDC hdc) {
    wchar_t buf[160];

    RECT trc = { 32, 22, CLIENT_W - 32, 58 };
    HGDIOBJ of = SelectObject(hdc, g_f_title);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, C_TEXT);
    DrawTextW(hdc, L"今天心情怎么样？", -1, &trc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, g_f_sub);
    SetTextColor(hdc, C_TEXT_SUB);
    RECT src = { 32, 60, CLIENT_W - 32, 84 };
    DrawTextW(hdc, L"选一个心情标签，再打个分", -1, &src, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);

    if (g_sel_tag >= 0 && g_sel_score > 0)
        wsprintfW(buf, L"%s · %d 分", kTags[g_sel_tag], g_sel_score);
    else if (g_sel_tag >= 0)
        wsprintfW(buf, L"#%s · 还没打分", kTags[g_sel_tag]);
    else if (g_sel_score > 0)
        wsprintfW(buf, L"%d 分 · 还没选心情", g_sel_score);
    else
        wsprintfW(buf, L"点上面的按钮开始记录");
    draw_text_c(hdc, buf, (g_sel_tag >= 0 || g_sel_score > 0) ? g_f_sel : g_f_sub,
                (g_sel_tag >= 0 && g_sel_score > 0) ? C_ACCENT : C_TEXT_SUB, 316);

    char d[16];
    today_str(d, sizeof d);
    wsprintfW(buf, L"今天已记 %d 条 · 累计 %d 条 · 连续记录 %d 天",
              count_day(d), g_ndates, calc_streak());
    draw_text_c(hdc, buf, g_f_status, C_TEXT_SUB, 346);
}

/* 全部按钮 + 文字区重绘 */
static void repaint_all(void) {
    for (int i = 0; i < kTagCount; i++) InvalidateRect(g_tagbtns[i], NULL, TRUE);
    for (int i = 0; i < 10; i++)        InvalidateRect(g_scorebtns[i], NULL, TRUE);
    RECT trc = { 0, 310, CLIENT_W, 384 };
    InvalidateRect(g_hwnd, &trc, TRUE);
}

/* ---------------- 界面 ---------------- */

static HWND mk_button(HWND parent, const wchar_t *text, int x, int y, int w, int ht, int id) {
    return CreateWindowExW(0, L"BUTTON", text,
                           WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP,
                           x, y, w, ht, parent, (HMENU)(INT_PTR)id, NULL, NULL);
}

static void on_create(HWND hwnd) {
    mkfonts();
    g_br_card = CreateSolidBrush(C_CARD);

    for (int i = 0; i < kTagCount; i++) {
        int col = i % 6, row = i / 6;
        g_tagbtns[i] = mk_button(hwnd, kTags[i],
                                 44 + col * 80, 114 + row * 52, 72, 40,
                                 IDC_TAG_BASE + i);
    }
    for (int i = 0; i < 10; i++) {
        wchar_t s[8];
        wsprintfW(s, L"%d", i + 1);
        g_scorebtns[i] = mk_button(hwnd, s, 43 + i * 48, 262, 42, 36,
                                   IDC_SCORE_BASE + i);
    }
    g_save = mk_button(hwnd, L"记下来", (CLIENT_W - 200) / 2, 390, 200, 48, IDC_SAVE);

    load_dates();
    repaint_all();
}

static void reset_selection(void) {
    g_sel_tag = -1;
    g_sel_score = 0;
}

static void on_save(HWND hwnd) {
    if (g_sel_tag < 0 && g_sel_score == 0) {
        MessageBoxW(hwnd, L"先挑一个心情，再打个分吧。", APP_TITLE, MB_OK | MB_ICONINFORMATION);
        return;
    }
    if (g_sel_tag < 0) {
        MessageBoxW(hwnd, L"还差一个心情标签。", APP_TITLE, MB_OK | MB_ICONINFORMATION);
        return;
    }
    if (g_sel_score == 0) {
        MessageBoxW(hwnd, L"还差一个分数（1~10）。", APP_TITLE, MB_OK | MB_ICONINFORMATION);
        return;
    }

    if (!append_record(g_sel_tag, g_sel_score)) {
        MessageBoxW(hwnd, L"写入文件失败，记录没有保存！", APP_TITLE, MB_OK | MB_ICONERROR);
        return;
    }

    load_dates();                     /* 重新统计 */

    wchar_t buf[160];
    char d[16];
    today_str(d, sizeof d);
    wsprintfW(buf, L"已经记下啦：%s，%d 分\n\n今天共 %d 条 · 连续记录 %d 天",
              kTags[g_sel_tag], g_sel_score, count_day(d), calc_streak());
    MessageBoxW(hwnd, buf, APP_TITLE, MB_OK | MB_ICONASTERISK);

    reset_selection();
    repaint_all();
}

static void on_command(HWND hwnd, int id, int code) {
    if (code != BN_CLICKED) return;

    if (id >= IDC_TAG_BASE && id < IDC_TAG_BASE + kTagCount) {
        g_sel_tag = id - IDC_TAG_BASE;
        repaint_all();
    } else if (id >= IDC_SCORE_BASE && id < IDC_SCORE_BASE + 10) {
        g_sel_score = id - IDC_SCORE_BASE + 1;
        repaint_all();
    } else if (id == IDC_SAVE) {
        on_save(hwnd);
    }
}

/* ---------------- 消息处理 ---------------- */

static BOOL on_drawitem(HWND hwnd, WPARAM wParam, LPARAM lParam) {
    (void)hwnd; (void)wParam;
    DRAWITEMSTRUCT *dis = (DRAWITEMSTRUCT *)lParam;
    int id = (int)dis->CtlID;
    int pressed = (dis->itemState & ODS_SELECTED) != 0;

    if (id >= IDC_TAG_BASE && id < IDC_TAG_BASE + kTagCount) {
        int sel = (g_sel_tag == id - IDC_TAG_BASE);
        COLORREF bg = sel ? C_ACCENT : (pressed ? C_CHIP_DN : C_CHIP_BG);
        draw_capsule(dis->hDC, dis->rcItem, bg, sel ? C_ACCENT : bg,
                     kTags[id - IDC_TAG_BASE], sel ? g_f_chip_b : g_f_chip,
                     sel ? RGB(255,255,255) : C_CHIP_TX);
        return TRUE;
    }
    if (id >= IDC_SCORE_BASE && id < IDC_SCORE_BASE + 10) {
        int sel = (g_sel_score == id - IDC_SCORE_BASE + 1);
        wchar_t s[8];
        wsprintfW(s, L"%d", id - IDC_SCORE_BASE + 1);
        COLORREF bg = sel ? C_ACCENT : (pressed ? C_CHIP_DN : C_CHIP_BG);
        draw_capsule(dis->hDC, dis->rcItem, bg, sel ? C_ACCENT : bg,
                     s, sel ? g_f_chip_b : g_f_chip,
                     sel ? RGB(255,255,255) : C_CHIP_TX);
        return TRUE;
    }
    if (id == IDC_SAVE) {
        draw_capsule(dis->hDC, dis->rcItem, pressed ? C_ACCENT_DK : C_ACCENT,
                     pressed ? C_ACCENT_DK : C_ACCENT, L"记下来", g_f_btn,
                     RGB(255,255,255));
        return TRUE;
    }
    return FALSE;
}

static LRESULT CALLBACK wnd_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        g_hwnd = hwnd;
        on_create(hwnd);
        return 0;

    case WM_COMMAND:
        on_command(hwnd, LOWORD(wParam), HIWORD(wParam));
        return 0;

    case WM_DRAWITEM:
        return on_drawitem(hwnd, wParam, lParam);

    case WM_ERASEBKGND:
        return 1;                       /* 背景在 WM_PAINT 双缓冲里画，防闪烁 */

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC wdc = BeginPaint(hwnd, &ps);
        HDC mem = CreateCompatibleDC(wdc);
        HBITMAP bmp = CreateCompatibleBitmap(wdc, CLIENT_W, CLIENT_H);
        HGDIOBJ ob = SelectObject(mem, bmp);
        paint_background(mem);
        paint_texts(mem);
        BitBlt(wdc, 0, 0, CLIENT_W, CLIENT_H, mem, 0, 0, SRCCOPY);
        SelectObject(mem, ob);
        DeleteObject(bmp);
        DeleteDC(mem);
        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_CTLCOLORBTN:
        return (LRESULT)g_br_card;      /* 按钮底色与卡片一致，消除边角黑框 */

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmd, int nShow) {
    (void)hPrev; (void)cmd;

    WNDCLASSW wc = { 0 };
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = wnd_proc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.hIcon         = LoadIconW(NULL, (LPCWSTR)IDI_APPLICATION);
    wc.hbrBackground = NULL;
    wc.lpszClassName = L"MoodDiaryWnd";
    RegisterClassW(&wc);

    RECT rc = { 0, 0, CLIENT_W, CLIENT_H };
    DWORD style = (WS_OVERLAPPEDWINDOW & ~(DWORD)(WS_THICKFRAME | WS_MAXIMIZEBOX)) | WS_MINIMIZEBOX | WS_CLIPCHILDREN;
    AdjustWindowRectEx(&rc, style, FALSE, 0);

    HWND hwnd = CreateWindowExW(0, L"MoodDiaryWnd", APP_TITLE, style,
                                CW_USEDEFAULT, CW_USEDEFAULT,
                                rc.right - rc.left, rc.bottom - rc.top,
                                NULL, NULL, hInst, NULL);
    ShowWindow(hwnd, nShow);
    UpdateWindow(hwnd);

    MSG m;
    while (GetMessageW(&m, NULL, 0, 0)) {
        TranslateMessage(&m);
        DispatchMessageW(&m);
    }
    return (int)m.wParam;
}
