#define _POSIX_C_SOURCE 200809L
#include <assert.h>
#include <ctype.h>
#include <math.h>
#include <ncurses.h>
#include <stdlib.h>
#include <string.h>

#define MAXIN 128                     // max cell input length
#define NCOL 256                      // max number of columns (A..Z)
#define NROW 1024                     // max number of rows
#define GW 4                          // row-number gutter
int CW = 8;                           // column display width
enum { EMPTY, NUM, LABEL, FORMULA };  // cell types
struct cell {
  int type;
  float val;
  char text[MAXIN];  // raw user input
  int fmt;  // 0=general 'I'=integer 'D'=default '$'=dollar '%'=percent '*'=graph 'L'=left 'R'=right
};

struct grid {
  struct cell cells[NCOL][NROW];
  int cc, cr, vc, vr, tc, tr, fmt, dirty;
  const char* filename;
};

struct parser {
  const char *s, *p;
  struct grid* g;
};

float expr(struct parser* p);
int ref(const char* s, int* col, int* row);
int refabs(const char* s, int* col, int* row, int* absc, int* absr);
static char* col(int c);

struct cell* cell(struct grid* g, int c, int r) {
  return (c >= 0 && c < NCOL && r >= 0 && r < NROW) ? &g->cells[c][r] : NULL;
}

void recalc(struct grid* g) {
  for (int pass = 0; pass < 100; pass++) {
    int changed = 0;
    for (int r = 0; r < NROW; r++)
      for (int c = 0; c < NCOL; c++) {
        struct cell* cl = &g->cells[c][r];
        if (cl->type != FORMULA) continue;
        struct parser p = {cl->text, cl->text, g};
        float v = expr(&p);
        if (v != cl->val) changed = 1;
        cl->val = v;
      }
    if (!changed) break;
  }
}

// Emit a cell reference into buf, preserving $ markers.
static int emitref(char* buf, int bufsz, int rc, int rr, int ac, int ar) {
  int oi = 0;
  if (ac && oi < bufsz) buf[oi++] = '$';
  oi += snprintf(buf + oi, bufsz - oi, "%s", col(rc));
  if (ar && oi < bufsz) buf[oi++] = '$';
  oi += snprintf(buf + oi, bufsz - oi, "%d", rr + 1);
  return oi;
}

// Rewrite cell references in all formulas after swapping two adjacent rows or columns.
// axis='R': rows a and b swapped. axis='C': columns a and b swapped.
static void fixrefs(struct grid* g, int axis, int a, int b) {
  for (int r = 0; r < NROW; r++)
    for (int c = 0; c < NCOL; c++) {
      struct cell* cl = &g->cells[c][r];
      if (cl->type != FORMULA) continue;
      char out[MAXIN] = {0};
      int oi = 0, changed = 0;
      const char* s = cl->text;
      while (*s && oi < MAXIN - 8) {
        int rc, rr, ac, ar, n = refabs(s, &rc, &rr, &ac, &ar);
        if (n) {
          if (axis == 'R') {
            if (rr == a)
              rr = b, changed = 1;
            else if (rr == b)
              rr = a, changed = 1;
          } else {
            if (rc == a)
              rc = b, changed = 1;
            else if (rc == b)
              rc = a, changed = 1;
          }
          oi += emitref(out + oi, MAXIN - oi, rc, rr, ac, ar);
          s += n;
        } else {
          out[oi++] = *s++;
        }
      }
      out[oi] = '\0';
      if (changed) strncpy(cl->text, out, MAXIN - 1);
    }
}

// Shift cell references after insert/delete.
// axis='R' for row, 'C' for column. pos=where inserted/deleted. dir=+1 insert, -1 delete.
static void shiftrefs(struct grid* g, int axis, int pos, int dir) {
  for (int r = 0; r < NROW; r++)
    for (int c = 0; c < NCOL; c++) {
      struct cell* cl = &g->cells[c][r];
      if (cl->type != FORMULA) continue;
      char out[MAXIN] = {0};
      int oi = 0, changed = 0;
      const char* s = cl->text;
      while (*s && oi < MAXIN - 8) {
        int rc, rr, ac, ar, n = refabs(s, &rc, &rr, &ac, &ar);
        if (n) {
          if (axis == 'R') {
            if (dir > 0 && rr >= pos)
              rr++, changed = 1;
            else if (dir < 0 && rr > pos)
              rr--, changed = 1;
          } else {
            if (dir > 0 && rc >= pos)
              rc++, changed = 1;
            else if (dir < 0 && rc > pos)
              rc--, changed = 1;
          }
          oi += emitref(out + oi, MAXIN - oi, rc, rr, ac, ar);
          s += n;
        } else {
          out[oi++] = *s++;
        }
      }
      out[oi] = '\0';
      if (changed) strncpy(cl->text, out, MAXIN - 1);
    }
}

static void insertrow(struct grid* g, int at) {
  for (int c = 0; c < NCOL; c++)
    for (int r = NROW - 1; r > at; r--) g->cells[c][r] = g->cells[c][r - 1];
  for (int c = 0; c < NCOL; c++) g->cells[c][at] = (struct cell){0};
  shiftrefs(g, 'R', at, +1);
  g->dirty = 1;
}

static void insertcol(struct grid* g, int at) {
  for (int r = 0; r < NROW; r++)
    for (int c = NCOL - 1; c > at; c--) g->cells[c][r] = g->cells[c - 1][r];
  for (int r = 0; r < NROW; r++) g->cells[at][r] = (struct cell){0};
  shiftrefs(g, 'C', at, +1);
  g->dirty = 1;
}

static void deleterow(struct grid* g, int at) {
  shiftrefs(g, 'R', at, -1);
  for (int c = 0; c < NCOL; c++)
    for (int r = at; r < NROW - 1; r++) g->cells[c][r] = g->cells[c][r + 1];
  for (int c = 0; c < NCOL; c++) g->cells[c][NROW - 1] = (struct cell){0};
  g->dirty = 1;
}

static void deletecol(struct grid* g, int at) {
  shiftrefs(g, 'C', at, -1);
  for (int r = 0; r < NROW; r++)
    for (int c = at; c < NCOL - 1; c++) g->cells[c][r] = g->cells[c + 1][r];
  for (int r = 0; r < NROW; r++) g->cells[NCOL - 1][r] = (struct cell){0};
  g->dirty = 1;
}

void setcell(struct grid* g, int c, int r, const char* input) {
  struct cell* cl = cell(g, c, r);
  if (!cl) return;
  if (!*input) {
    *cl = (struct cell){0};
    recalc(g);
    return;
  }

  strncpy(cl->text, input, MAXIN - 1);
  g->dirty = 1;

  if (input[0] == '+' || input[0] == '-' || input[0] == '(' || input[0] == '@') {
    cl->type = FORMULA;
  } else if (isdigit(input[0]) || input[0] == '.') {
    char* end;
    double v = strtod(input, &end);
    cl->type = (*end == '\0') ? NUM : FORMULA;
    if (cl->type == NUM) cl->val = v;
  } else {
    cl->type = LABEL;
    cl->val = 0;
  }
  recalc(g);
}

void skipws(struct parser* p) { for (; isspace(*p->p); p->p++); }

float number(struct parser* p) {
  char* end;
  float v = strtof(p->p, &end);
  if (end == p->p) return NAN;
  p->p = end;
  return v;
}

int refabs(const char* s, int* col, int* row, int* absc, int* absr) {
  const char* p = s;
  *absc = *absr = 0;
  if (*p == '$') {
    *absc = 1;
    p++;
  }
  if (!isalpha(*p)) return 0;
  *col = toupper(*p++) - 'A' + 1;
  if (isalpha(*p)) *col = *col * 26 + (toupper(*p++) - 'A' + 1);
  if (*p == '$') {
    *absr = 1;
    p++;
  }
  char* end;
  int n = strtol(p, &end, 10);
  if (n <= 0 || end == p) return 0;
  *row = n - 1, *col = *col - 1;
  return (int)(end - s);
}

int ref(const char* s, int* col, int* row) {
  int ac, ar;
  return refabs(s, col, row, &ac, &ar);
}

static float cellval(struct parser* p) {
  int c, r, ac, ar, n = refabs(p->p, &c, &r, &ac, &ar);
  if (!n) return NAN;
  p->p += n;
  struct cell* cl = cell(p->g, c, r);
  if (!cl) return NAN;
  return cl->val;
}

float func(struct parser* p) {
  char fn[16] = {0};
  for (int i = 0; isalpha(*p->p) && i < sizeof(fn) - 1;) fn[i++] = toupper(*p->p++);

  if (*p->p != '(') return NAN;
  p->p++;
  skipws(p);

  // Try to parse range: A1...B5
  int c1, r1, c2, r2, is_range = 0;
  const char* q = p->p;
  int n = ref(q, &c1, &r1);
  if (n) {
    q += n;
    if (*q == '.' && *(q + 1) == '.' && *(q + 2) == '.') {
      q += 3;
      n = ref(q, &c2, &r2);
      if (n) {
        p->p = q + n;
        is_range = 1;
      }
    }
  }

  float result = 0;
  if (is_range && strcmp(fn, "SUM") == 0) {
    if (c1 > c2) n = c1, c1 = c2, c2 = n;
    if (r1 > r2) n = r1, r1 = r2, r2 = n;
    for (int c = c1; c <= c2; c++)
      for (int r = r1; r <= r2; r++) {
        struct cell* cl = cell(p->g, c, r);
        if (cl && cl->type != EMPTY && cl->type != LABEL) result += cl->val;
      }
  } else if (!is_range) {
    float arg = expr(p);
    if (strcmp(fn, "SUM") == 0)
      result = arg;
    else if (strcmp(fn, "ABS") == 0)
      result = fabs(arg);
    else if (strcmp(fn, "INT") == 0)
      result = (int)arg;
    else if (strcmp(fn, "SQRT") == 0)
      result = arg >= 0 ? sqrt(arg) : NAN;
    else
      return NAN;
  } else {
    return NAN;
  }

  skipws(p);
  if (*p->p != ')') return NAN;
  p->p++;
  return result;
}

float primary(struct parser* p) {
  skipws(p);
  if (!*p->p) return NAN;
  if (*p->p == '+') p->p++;
  if (*p->p == '-') {
    p->p++;
    return -primary(p);
  }
  if (*p->p == '@') {
    p->p++;
    return func(p);
  }
  if (*p->p == '(') {
    p->p++;
    float v = expr(p);
    skipws(p);
    if (*p->p != ')') return NAN;
    p->p++;
    return v;
  }
  if (isdigit(*p->p) || *p->p == '.') return number(p);
  return cellval(p);
}

float term(struct parser* p) {
  float l = primary(p);
  for (;;) {
    skipws(p);
    char op = *p->p;
    if (op != '*' && op != '/') break;
    p->p++;
    float r = primary(p);
    if (op == '*')
      l *= r;
    else if (r == 0)
      return NAN;
    else
      l /= r;
  }
  return l;
}

float expr(struct parser* p) {
  float l = term(p);
  for (;;) {
    skipws(p);
    char op = *p->p;
    if (op != '+' && op != '-') break;
    p->p++;
    float r = term(p);
    l = (op == '+') ? l + r : l - r;
  }
  return l;
}

static int csvfield(FILE* f, char* buf, int bufsz, int* eol, int* eof) {
  int c, n = 0, quoted = 0;
  *eol = *eof = 0;
  buf[0] = '\0';

  c = fgetc(f);
  if (c == EOF) {
    *eof = 1;
    return 0;
  }

  if (c == '"') {
    quoted = 1;
    for (;;) {
      c = fgetc(f);
      if (c == EOF) {
        *eof = 1;
        break;
      }
      if (c == '"') {
        c = fgetc(f);
        if (c == '"') {
          if (n < bufsz - 1) buf[n++] = '"';
        } else {
          if (c == '\r') c = fgetc(f);
          if (c == '\n' || c == EOF) {
            *eol = 1;
            if (c == EOF) *eof = 1;
          }
          break;
        }
      } else {
        if (n < bufsz - 1) buf[n++] = c;
      }
    }
  } else {
    for (;;) {
      if (c == ',' || c == '\n' || c == EOF) {
        if (c == '\n' || c == EOF) {
          *eol = 1;
          if (c == EOF) *eof = 1;
        }
        break;
      }
      if (c == '\r') {
        c = fgetc(f);
        continue;
      }
      if (n < bufsz - 1) buf[n++] = c;
      c = fgetc(f);
    }
  }
  buf[n] = '\0';
  return 1;
}

int csvload(struct grid* g, const char* filename) {
  FILE* f = fopen(filename, "r");
  if (!f) return -1;
  char buf[MAXIN];
  int row = 0, col = 0, eol, eof;

  while (csvfield(f, buf, sizeof(buf), &eol, &eof)) {
    if (buf[0] && row < NROW && col < NCOL) setcell(g, col, row, buf);
    if (eol) {
      row++, col = 0;
    } else
      col++;
    if (eof) break;
  }
  fclose(f);
  return 0;
}

static int csvneedsquote(const char* s) {
  for (; *s; s++)
    if (*s == ',' || *s == '"' || *s == '\n' || *s == '\r') return 1;
  return 0;
}

static void csvwritefield(FILE* f, const char* s) {
  if (csvneedsquote(s)) {
    fputc('"', f);
    for (; *s; s++) {
      if (*s == '"') fputc('"', f);
      fputc(*s, f);
    }
    fputc('"', f);
  } else {
    fputs(s, f);
  }
}

int csvsave(struct grid* g, const char* filename) {
  FILE* f = fopen(filename, "w");
  if (!f) return -1;

  int maxr = -1, maxc = -1;
  for (int r = 0; r < NROW; r++)
    for (int c = 0; c < NCOL; c++)
      if (g->cells[c][r].type != EMPTY) {
        if (r > maxr) maxr = r;
        if (c > maxc) maxc = c;
      }

  for (int r = 0; r <= maxr; r++) {
    for (int c = 0; c <= maxc; c++) {
      if (c > 0) fputc(',', f);
      struct cell* cl = &g->cells[c][r];
      if (cl->type == EMPTY) continue;
      csvwritefield(f, cl->text);
    }
    fputc('\n', f);
  }
  fclose(f);
  return 0;
}

static int vcols(void) { return (COLS - GW) / CW > 0 ? (COLS - GW) / CW : 1; }
static int vrows(void) { return (LINES - 4) > 0 ? (LINES - 4) : 1; }

static char* col(int c) {
  static char buf[16] = {0};
  buf[0] = (c < 26) ? ('A' + c) : ('A' + (c) / 26 - 1);
  buf[1] = (c < 26) ? '\0' : ('A' + (c) % 26);
  return buf;
}

static void fmtcell(struct grid* g, struct cell* cl, char* fb, int cw) {
  if (!cl || cl->type == EMPTY) {
    memset(fb, ' ', cw);
    fb[cw] = '\0';
  } else if (cl->type == LABEL) {
    snprintf(fb, cw + 1, "%-*.*s", cw, cw, cl->text[0] == '"' ? cl->text + 1 : cl->text);
  } else if (isnan(cl->val)) {
    snprintf(fb, cw + 1, "%*s", cw, "ERROR");
  } else {
    char t[MAXIN] = {0}, fmt = cl->fmt;
    if (!fmt || fmt == 'D') fmt = g->fmt;
    if (fmt == '$') {
      snprintf(t, sizeof(t), "%.2f", cl->val);
    } else if (fmt == '%') {
      snprintf(t, sizeof(t), "%.2f%%", cl->val * 100);
    } else if (fmt == '*') {
      for (int i = 0; i < cw && i < cl->val; i++) t[i] = '*';
      fmt = 'L';
    } else if (fmt == 'I' || (cl->val == (long)cl->val && fabs(cl->val) < 1e9)) {
      snprintf(t, sizeof(t), "%ld", (long)cl->val);
    } else {
      snprintf(t, sizeof(t), "%g", cl->val);
    }
    snprintf(fb, cw + 1, fmt == 'L' ? "%-*s" : "%*s", cw, t);
  }
}

static void draw(struct grid* g, const char* mode, const char* buf) {
  erase();

  int lc = g->tc;         // locked title columns
  int lr = g->tr;         // locked title rows
  int fc = vcols() - lc;  // free scrollable column slots
  int fr = vrows() - lr;  // free scrollable row slots
  if (fc < 1) fc = 1;
  if (fr < 1) fr = 1;

  // Status bar
  attron(A_BOLD | A_REVERSE);
  move(0, 0);
  clrtoeol();
  struct cell* cur = cell(g, g->cc, g->cr);
  mvprintw(0, 0, " %s%d", col(g->cc), g->cr + 1);
  if (cur && cur->type == NUM)
    printw("  %.10g", cur->val);
  else if (cur && cur->type == FORMULA)
    printw("  %s = %s%.10g", cur->text, isnan(cur->val) ? "ERR " : "",
           isnan(cur->val) ? 0.0 : cur->val);
  else if (cur && cur->type == LABEL)
    printw("  %s", cur->text);
  mvprintw(0, COLS - 6, "%s", mode);
  attroff(A_BOLD | A_REVERSE);

  // Edit line
  move(1, 0);
  clrtoeol();
  if (mode)
    mvprintw(1, 0, "%s_", buf);
  else if (cur && cur->type != EMPTY)
    mvprintw(1, 0, "  %s", cur->text);

  /* Row 2: column headers */
  attron(A_BOLD | A_REVERSE);
  move(2, 0);
  clrtoeol();
  for (int ci = 0; ci < lc + fc; ci++) {
    int c = (ci < lc) ? ci : g->vc + (ci - lc);
    if (c >= NCOL) break;
    mvprintw(2, GW + ci * CW, "%*s", CW, col(c));
  }
  attroff(A_BOLD | A_REVERSE);

  // Grid rows: locked title rows first, then scrollable rows
  for (int ri = 0; ri < lr + fr; ri++) {
    int row = (ri < lr) ? ri : g->vr + (ri - lr);
    if (row >= NROW) continue;
    int y = 3 + ri;
    int is_locked_row = (ri < lr);

    move(y, 0);
    clrtoeol();
    attron(A_REVERSE);
    mvprintw(y, 0, "%*d ", GW - 1, row + 1);
    attroff(A_REVERSE);

    for (int ci = 0; ci < lc + fc; ci++) {
      int c = (ci < lc) ? ci : g->vc + (ci - lc);
      if (c >= NCOL) break;
      int is_locked_col = (ci < lc);

      struct cell* cl = cell(g, c, row);
      char fb[64];
      fmtcell(g, cl, fb, CW);

      int is_cur = (c == g->cc && row == g->cr);
      int is_locked = (is_locked_row || is_locked_col);
      if (is_cur || is_locked) attron(A_REVERSE);
      if (is_locked && !is_cur) attron(A_BOLD);
      mvprintw(y, GW + ci * CW, "%s", fb);
      if (is_locked && !is_cur) attroff(A_BOLD);
      if (is_cur || is_locked) attroff(A_REVERSE);
    }
  }
}

static void swaprow(struct grid* g, int a, int b) {
  for (int c = 0; c < NCOL; c++) {
    struct cell tmp = g->cells[c][a];
    g->cells[c][a] = g->cells[c][b];
    g->cells[c][b] = tmp;
  }
  fixrefs(g, 'R', a, b);
}

static void swapcol(struct grid* g, int a, int b) {
  for (int r = 0; r < NROW; r++) {
    struct cell tmp = g->cells[a][r];
    g->cells[a][r] = g->cells[b][r];
    g->cells[b][r] = tmp;
  }
  fixrefs(g, 'C', a, b);
}

void movecmd(struct grid* g) {
  int origc = g->cc, origr = g->cr;
  char src[16];
  snprintf(src, sizeof(src), "%s%d", col(origc), origr + 1);
  for (;;) {
    draw(g, "MOVE", "");
    if (g->cc == origc && g->cr == origr)
      mvprintw(1, 0, "Source: %s  (move cursor, Esc cancel)", src);
    else
      mvprintw(1, 0, "%s...%s%d  (Enter confirm, Esc cancel)", src, col(g->cc), g->cr + 1);
    clrtoeol();
    refresh();
    int k = getch();
    if (k == 27) {  // cancel: undo all swaps
      if (g->cc != origc) {
        while (g->cc < origc) swapcol(g, g->cc, g->cc + 1), g->cc++;
        while (g->cc > origc) swapcol(g, g->cc, g->cc - 1), g->cc--;
      } else {
        while (g->cr < origr) swaprow(g, g->cr, g->cr + 1), g->cr++;
        while (g->cr > origr) swaprow(g, g->cr, g->cr - 1), g->cr--;
      }
      recalc(g);
      break;
    } else if (k == 10 || k == 13 || k == KEY_ENTER) {
      if (g->cc != origc || g->cr != origr) g->dirty = 1;
      recalc(g);
      break;
    } else if (k == KEY_UP && g->cc == origc) {
      int lo = g->tr > 0 ? g->tr : 0;
      if (g->cr > lo) {
        swaprow(g, g->cr, g->cr - 1);
        g->cr--;
      }
    } else if (k == KEY_DOWN && g->cc == origc) {
      if (g->cr < NROW - 1) {
        swaprow(g, g->cr, g->cr + 1);
        g->cr++;
      }
    } else if (k == KEY_LEFT && g->cr == origr) {
      int lo = g->tc > 0 ? g->tc : 0;
      if (g->cc > lo) {
        swapcol(g, g->cc, g->cc - 1);
        g->cc--;
      }
    } else if (k == KEY_RIGHT && g->cr == origr) {
      if (g->cc < NCOL - 1) {
        swapcol(g, g->cc, g->cc + 1);
        g->cc++;
      }
    }
  }
}

static void replicatecell(struct grid* g, int sc, int sr, int dc, int dr) {
  struct cell* src = cell(g, sc, sr);
  struct cell* dst = cell(g, dc, dr);
  if (!src || !dst) return;
  if (src->type == EMPTY) {
    *dst = (struct cell){0};
    return;
  }
  *dst = *src;
  if (src->type != FORMULA) return;

  int dcol = dc - sc, drow = dr - sr;
  char out[MAXIN] = {0};
  int oi = 0;
  const char* s = src->text;
  while (*s && oi < MAXIN - 8) {
    int rc, rr, ac, ar, n = refabs(s, &rc, &rr, &ac, &ar);
    if (n) {
      if (!ac) rc += dcol;
      if (!ar) rr += drow;
      oi += emitref(out + oi, MAXIN - oi, rc, rr, ac, ar);
      s += n;
    } else {
      out[oi++] = *s++;
    }
  }
  out[oi] = '\0';
  strncpy(dst->text, out, MAXIN - 1);
}

// Helper: format a range string into buf (uses static col() buffer carefully)
static void fmtrange(char* buf, int sz, int c1, int r1, int c2, int r2) {
  if (c1 == c2 && r1 == r2) {
    snprintf(buf, sz, "%s%d", col(c1), r1 + 1);
  } else {
    char a[16];
    snprintf(a, sizeof(a), "%s%d", col(c1), r1 + 1);
    snprintf(buf, sz, "%s...%s%d", a, col(c2), r2 + 1);
  }
}

// Select a range: anchor is fixed, cursor moves to define other corner.
// Both cursor keys and typed input work. Returns 1 on Enter, 0 on Esc.
static int selectrange(struct grid* g, const char* prompt, int ac, int ar, int* c1, int* r1,
                       int* c2, int* r2) {
  char buf[MAXIN] = {0};
  int n = 0, typed = 0;
  g->cc = ac;
  g->cr = ar;
  for (;;) {
    char rng[MAXIN + 4];
    if (typed) {
      snprintf(rng, sizeof(rng), "%s_", buf);
    } else {
      fmtrange(rng, sizeof(rng), ac < g->cc ? ac : g->cc, ar < g->cr ? ar : g->cr,
               ac > g->cc ? ac : g->cc, ar > g->cr ? ar : g->cr);
    }
    draw(g, "REPL", "");
    mvprintw(1, 0, "%s %s", prompt, rng);
    clrtoeol();
    refresh();
    int ch = getch();
    if (ch == 27) return 0;
    if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
      if (typed) {
        const char* p = buf;
        int k = ref(p, c1, r1);
        if (!k) return 0;
        p += k;
        *c2 = *c1, *r2 = *r1;
        if (p[0] == '.' && p[1] == '.' && p[2] == '.') {
          p += 3;
          if (!ref(p, c2, r2)) return 0;
        }
      } else {
        *c1 = ac < g->cc ? ac : g->cc;
        *r1 = ar < g->cr ? ar : g->cr;
        *c2 = ac > g->cc ? ac : g->cc;
        *r2 = ar > g->cr ? ar : g->cr;
      }
      if (*c1 > *c2) {
        int t = *c1;
        *c1 = *c2;
        *c2 = t;
      }
      if (*r1 > *r2) {
        int t = *r1;
        *r1 = *r2;
        *r2 = t;
      }
      return 1;
    } else if (ch == KEY_UP || ch == KEY_DOWN || ch == KEY_LEFT || ch == KEY_RIGHT) {
      typed = 0;
      buf[0] = '\0';
      n = 0;
      if (ch == KEY_UP && g->cr > 0)
        g->cr--;
      else if (ch == KEY_DOWN && g->cr < NROW - 1)
        g->cr++;
      else if (ch == KEY_LEFT && g->cc > 0)
        g->cc--;
      else if (ch == KEY_RIGHT && g->cc < NCOL - 1)
        g->cc++;
    } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
      typed = 1;
      if (n > 0) buf[--n] = '\0';
    } else if (n < MAXIN - 1 && ch >= 32 && ch < 127) {
      typed = 1;
      buf[n++] = toupper(ch);
      buf[n] = '\0';
    }
  }
}

void replcmd(struct grid* g) {
  int sc1, sr1, sc2, sr2;
  int origc = g->cc, origr = g->cr;

  // Phase 1: select source range (anchor = current cell)
  if (!selectrange(g, "Source:", origc, origr, &sc1, &sr1, &sc2, &sr2)) return;

  // Phase 2: pick target top-left corner, target size = source size
  int sw = sc2 - sc1 + 1, sh = sr2 - sr1 + 1;
  char srcstr[32];
  fmtrange(srcstr, sizeof(srcstr), sc1, sr1, sc2, sr2);
  g->cc = sc1, g->cr = sr1;  // start target selection from source position
  char buf[MAXIN] = {0};
  int n = 0, typed = 0;
  for (;;) {
    char tgt[MAXIN + 4];
    int tc = typed ? -1 : g->cc, tr = typed ? -1 : g->cr;
    if (typed) {
      snprintf(tgt, sizeof(tgt), "%s_", buf);
    } else {
      fmtrange(tgt, sizeof(tgt), tc, tr, tc + sw - 1, tr + sh - 1);
    }
    draw(g, "REPL", "");
    mvprintw(1, 0, "%s to: %s", srcstr, tgt);
    clrtoeol();
    refresh();
    int ch = getch();
    if (ch == 27) return;
    if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
      int tc1, tr1;
      if (typed) {
        int dummy;
        int k = ref(buf, &tc1, &tr1);
        if (!k) return;
      } else {
        tc1 = g->cc;
        tr1 = g->cr;
      }
      for (int r = 0; r < sh; r++)
        for (int c = 0; c < sw; c++) replicatecell(g, sc1 + c, sr1 + r, tc1 + c, tr1 + r);
      recalc(g);
      g->dirty = 1;
      return;
    } else if (ch == KEY_UP || ch == KEY_DOWN || ch == KEY_LEFT || ch == KEY_RIGHT) {
      typed = 0;
      buf[0] = '\0';
      n = 0;
      if (ch == KEY_UP && g->cr > 0)
        g->cr--;
      else if (ch == KEY_DOWN && g->cr < NROW - 1)
        g->cr++;
      else if (ch == KEY_LEFT && g->cc > 0)
        g->cc--;
      else if (ch == KEY_RIGHT && g->cc < NCOL - 1)
        g->cc++;
    } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
      typed = 1;
      if (n > 0) buf[--n] = '\0';
    } else if (n < MAXIN - 1 && ch >= 32 && ch < 127) {
      typed = 1;
      buf[n++] = toupper(ch);
      buf[n] = '\0';
    }
  }
}

//
//  /B                   Blank current cell value (keep formatting)
//  /C                   Clear entire spreadsheet (keep formatting)
//  /F(L/R/I/G/D/$/%/*)  Set cell format: Left/Right/Integer/General/Dollar/Percent
//  /DR, /DC             Delete row/column
//  /IR, /IC             Insert row/column
//  /GC                  Set column width
//  /GF(L/R/I/G/D/$/%/*) Set default column format
//  /M                   Move row/column
//  /R                   Replicate cell
//  /SL                  Load CSV file
//  /SS                  Save CSV file
//  /SQ                  Save and quit
//  /T(V/H/B/N)          Lock rows/columns
//  /Q                   Quit (prompts if unsaved)
//
int command(struct grid* g) {
  draw(g, "CMD", "");
  mvprintw(1, 0, "Command: B C F D I G M R S T Q"), clrtoeol();
  refresh();
  int ch = toupper(getch());
  if (ch == 'B') {  // blank current cell
    setcell(g, g->cc, g->cr, "");
    recalc(g);
  } else if (ch == 'C') {  // clear sheet
    mvprintw(1, 0, "Clear entire sheet? (y/N)"), clrtoeol();
    ch = getch();
    if (ch == 'y' || ch == 'Y') {
      for (int r = 0; r < NROW; r++)
        for (int c = 0; c < NCOL; c++) g->cells[c][r] = (struct cell){0};
      g->dirty = 1;
    }
  } else if (ch == 'D') {  // delete row/column
    mvprintw(1, 0, "Delete (R)ow or (C)olumn?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'R')
      deleterow(g, g->cr);
    else if (ch == 'C')
      deletecol(g, g->cc);
    recalc(g);
  } else if (ch == 'I') {  // insert row/column
    mvprintw(1, 0, "Insert (R)ow or (C)olumn?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'R')
      insertrow(g, g->cr);
    else if (ch == 'C')
      insertcol(g, g->cc);
    recalc(g);
  } else if (ch == 'F') {  // change cell format
    mvprintw(1, 0, "Format: L R I G D $ %% *"), clrtoeol();
    ch = toupper(getch());
    struct cell* cl = cell(g, g->cc, g->cr);
    if (strchr("LRIGD$%*", ch)) cl->fmt = ch;
  } else if (ch == 'G') {  // change global settings
    mvprintw(1, 0, "Global: (C)ol width or (F)mt?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'C') {
      mvprintw(1, 0, "New column width (4-20)?"), clrtoeol();
      char buf[16] = {0};
      int n = 0;
      for (;;) {
        mvprintw(1, 30, "%s_", buf);
        int ch = getch();
        if (ch == 27) break;
        if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
          int w = strtol(buf, NULL, 10);
          if (w >= 4 && w <= 20) CW = w;
          break;
        } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
          if (n > 0) buf[--n] = '\0';
        } else if (isdigit(ch) && n < sizeof(buf) - 1) {
          buf[n++] = ch;
          buf[n] = '\0';
        }
      }
    } else if (ch == 'F') {
      mvprintw(1, 0, "Format: L R I G D $ %% *"), clrtoeol();
      ch = toupper(getch());
      if (strchr("LRIGD$%*", ch)) g->fmt = ch;
    }
  } else if (ch == 'M') {
    movecmd(g);
  } else if (ch == 'R') {
    replcmd(g);
  } else if (ch == 'T') {
    mvprintw(1, 0, "Lock (V)ertical, (H)orizontal, (B)oth, or (N)one?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'V') {
      g->tc = g->cc + 1, g->tr = 0;
      g->cc++;  // move cursor past locked columns
    } else if (ch == 'H') {
      g->tr = g->cr + 1, g->tc = 0;
      g->cr++;  // move cursor past locked rows
    } else if (ch == 'B') {
      g->tc = g->cc + 1, g->tr = g->cr + 1;
      g->cc++;
      g->cr++;
    } else if (ch == 'N') {
      g->tc = g->tr = 0;
      g->vc = g->vr = 0;
    }
  } else if (ch == 'S') {
    mvprintw(1, 0, "Storage: (L)oad, (S)ave, or save & (Q)uit?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'L') {
      // prompt for filename
      mvprintw(1, 0, "Load file: "), clrtoeol();
      char fbuf[256] = {0};
      int fn = 0;
      if (g->filename) {
        strncpy(fbuf, g->filename, sizeof(fbuf) - 1);
        fn = strlen(fbuf);
      }
      for (;;) {
        mvprintw(1, 11, "%s_  ", fbuf);
        int ch = getch();
        if (ch == 27) break;
        if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
          if (fn > 0) {
            for (int r = 0; r < NROW; r++)
              for (int c = 0; c < NCOL; c++) g->cells[c][r] = (struct cell){0};
            if (csvload(g, fbuf) == 0) {
              g->filename = strdup(fbuf);
              g->dirty = 0;
            } else {
              mvprintw(1, 0, "Failed to load: %s. Press any key.", fbuf), clrtoeol();
              getch();
            }
          }
          break;
        } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
          if (fn > 0) fbuf[--fn] = '\0';
        } else if (fn < (int)sizeof(fbuf) - 1 && ch >= 32) {
          fbuf[fn++] = ch;
          fbuf[fn] = '\0';
        }
      }
    } else if (ch == 'S' || ch == 'Q') {
      int quit = (ch == 'Q');
      const char* fn = g->filename;
      if (!fn) {
        // prompt for filename
        mvprintw(1, 0, "Save as: "), clrtoeol();
        static char sbuf[256] = {0};
        int sn = 0;
        sbuf[0] = '\0';
        for (;;) {
          mvprintw(1, 9, "%s_  ", sbuf);
          int ch = getch();
          if (ch == 27) {
            fn = NULL;
            break;
          }
          if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
            fn = (sn > 0) ? sbuf : NULL;
            break;
          } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
            if (sn > 0) sbuf[--sn] = '\0';
          } else if (sn < (int)sizeof(sbuf) - 1 && ch >= 32) {
            sbuf[sn++] = ch;
            sbuf[sn] = '\0';
          }
        }
      }
      if (fn) {
        if (csvsave(g, fn) == 0) {
          g->filename = fn;
          g->dirty = 0;
          if (quit) return 1;
        } else {
          mvprintw(1, 0, "Failed to save: %s. Press any key.", fn), clrtoeol();
          getch();
        }
      }
    }
  } else if (ch == 'Q') {
    if (g->dirty) {
      mvprintw(1, 0, "Unsaved changes. Quit anyway? (y/N)"), clrtoeol();
      ch = getch();
      if (ch == 'y' || ch == 'Y') return 1;
    } else {
      return 1;
    }
  }
  return 0;
}

// navigation mode: enter cell reference to jump to (e.g. B12)
void nav(struct grid* g) {
  char buf[MAXIN] = {0}, n = 0;
  draw(g, "GOTO", "");
  for (;;) {
    mvprintw(1, 0, "> %s_", buf);
    clrtoeol();
    int ch = getch();
    if (ch == 27) break;
    if (ch == 10 || ch == 13 || ch == KEY_ENTER || ch == 9) {
      int c, r;
      if (ref(buf, &c, &r)) g->cc = c, g->cr = r;
      break;
    } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
      if (n > 0) buf[--n] = '\0';
    } else if ((isalpha(ch) || isdigit(ch)) && strlen(buf) < MAXIN - 2) {
      int c, r, i = n;
      buf[i++] = toupper(ch);
      if (isalpha(ch)) buf[i++] = '1';
      if (ref(buf, &c, &r) && r < NROW && c < NCOL) n++;
      buf[n] = '\0';
    }
  }
}

// entry mode: edit cell content, label mode if label=1 or ch is non-formula starter
void entry(struct grid* g, int label, int ch) {
  char buf[MAXIN] = {0};
  int n;
  draw(g, "ENTRY", "");
  if (ch) buf[n++] = ch;
  for (;;) {
    mvprintw(1, 0, "> %s_", buf);
    clrtoeol();
    int ch = getch();
    if (ch == 27) break;
    if (ch == 10 || ch == 13 || ch == KEY_ENTER) {
      setcell(g, g->cc, g->cr, buf);
      if (g->cr < NROW - 1) g->cr++;
      break;
    } else if (ch == 9) {
      setcell(g, g->cc, g->cr, buf);
      if (g->cc < NCOL - 1) g->cc++;
      break;
    } else if (ch == KEY_BACKSPACE || ch == 127 || ch == 8) {
      if (n > 0) buf[--n] = '\0';
    } else if (n < MAXIN - 1) {
      buf[n++] = ch;
      buf[n] = '\0';
    }
  }
}

void loop(struct grid* g) {
  for (;;) {
    int lc = g->tc;
    int lr = g->tr;
    int fc = vcols() - lc;
    if (fc < 1) fc = 1;
    int fr = vrows() - lr;
    if (fr < 1) fr = 1;

    // Clamp cursor out of locked title area
    if (lc > 0 && g->cc < lc) g->cc = lc;
    if (lr > 0 && g->cr < lr) g->cr = lr;

    // Viewport must not overlap with locked area
    if (lc > 0 && g->vc < lc) g->vc = lc;
    if (lr > 0 && g->vr < lr) g->vr = lr;

    // Keep cursor visible in scrollable region
    if (g->cc >= lc) {
      if (g->cc < g->vc) g->vc = g->cc;
      if (g->cc >= g->vc + fc) g->vc = g->cc - fc + 1;
    }
    if (g->cr >= lr) {
      if (g->cr < g->vr) g->vr = g->cr;
      if (g->cr >= g->vr + fr) g->vr = g->cr - fr + 1;
    }
    draw(g, "READY", "");
    int ch = getch();

    if (ch == (0x1f & 'c'))  // Ctrl+C: quit
      break;
    else if (ch == KEY_UP && g->cr > lr)
      g->cr--;
    else if (ch == KEY_DOWN && g->cr < NROW - 1)
      g->cr++;
    else if (ch == KEY_LEFT && g->cc > lc)
      g->cc--;
    else if (ch == KEY_RIGHT && g->cc < NCOL - 1)
      g->cc++;
    else if (ch == KEY_HOME) {
      g->cc = lc;
      g->cr = lr;
    } else if (ch == 9 && g->cc < NCOL - 1)  // Tab: next cell
      g->cc++;
    else if (ch == 10 || ch == 13 || ch == KEY_ENTER) {  // Enter: next row
      if (g->cr < NROW - 1) g->cr++;
    } else if (ch == 127 || ch == 8 || ch == KEY_BACKSPACE) {  // Bsp/Del: clear cell
      struct cell* cl = cell(g, g->cc, g->cr);
      if (cl) *cl = (struct cell){0};
      recalc(g);
    } else if (ch == '!') {  // Recalculate
      recalc(g);
    } else if (ch == '/') {
      if (command(g)) break;
    } else if (ch == '>') {
      nav(g);
    } else if (ch == '"') {
      entry(g, 1, 0);
    } else if (ch == '+' || ch == '-' || ch == '(' || ch == '@' || ch == '.' || isdigit(ch)) {
      entry(g, 0, ch);
    } else if (ch >= 32 && ch < 127) {
      entry(g, 1, ch);
    }
  }
}

#ifndef TEST
int main(int argc, char* argv[]) {
  static struct grid g = {0};  // might be too large for stack
  if (argc == 2 && (strcmp(argv[1], "-h") == 0 || strcmp(argv[1], "--help") == 0)) {
    fprintf(stderr, "Usage: %s sheet.csv\n", argv[0]);
    exit(1);
  }
  if (argc > 1) {
    if (csvload(&g, argv[1]) < 0) {
      perror("Failed to load CSV file");
      exit(1);
    }
    g.filename = argv[1];
  }
  initscr();
  raw();
  keypad(stdscr, TRUE);
  noecho();
  curs_set(0);
  loop(&g);
  endwin();
  return 0;
}
#else
void test_expr(void) {
  struct grid g = {0};
  struct parser p = {0};
  g.cells[0][0].type = NUM;
  g.cells[0][0].val = 3.0f;
  g.cells[0][1].type = NUM;
  g.cells[0][1].val = 5.0f;
  g.cells[0][2].type = NUM;
  g.cells[0][2].val = 11.0f;
  g.cells[0][3].type = NUM;
  g.cells[0][3].val = -13.5f;
#define EXPR(e) (p.s = p.p = e, p.g = &g, expr(&p))
  // numbers
  assert(isnan(EXPR("")));
  assert(EXPR("42") == 42.0f);
  assert(EXPR("1.5") == 1.5f);
  assert(EXPR(".5") == 0.5f);
  assert(EXPR("-123") == -123.0f);
  assert(EXPR("+123") == 123.0f);
  assert(EXPR("(123)") == 123.0f);
  // references
  assert(EXPR("A1") == 3.0f);
  assert(EXPR("A2") == 5.0f);
  assert(EXPR("A3") == 11.0f);
  assert(EXPR("B1") == 0.0f);
  assert(EXPR("A12") == 0.0f);
  // expressions
  assert(EXPR("A1*A2") == 15.0f);
  assert(EXPR("A1*10/A2") == 6.0f);
  assert(isnan(EXPR("A1/0")));
  assert(EXPR("A1+A2") == 8.0f);
  assert(EXPR("A1+A2-A3") == -3.0f);
  assert(EXPR("A1+A2*A3") == 58.0f);
  assert(EXPR("(A1+A2)*A3") == 88.0f);
  // functions
  assert(EXPR("@ABS(A1)") == 3.0f);
  assert(EXPR("@ABS(A4)") == 13.5f);
  assert(EXPR("@INT(A4)") == -13.0f);
  assert(EXPR("@INT(@ABS(A4))") == 13.0f);
  assert(EXPR("@SQRT(A3+A2))") == 4.0f);
  assert(EXPR("@SUM(A3))") == 11.0f);
  assert(EXPR("@SUM(A1...A3))") == 19.0f);
  assert(EXPR("@SUM(A3...A1))") == 19.0f);
  assert(EXPR("@SUM(A1...A1))") == 3.0f);
  assert(EXPR("@SUM(A3...Z1))") == 19.0f);
#undef EXPR
}

void test_recalc(void) {
  struct grid g = {0};
  // testing re-evaluation
  setcell(&g, 0, 0, "5");
  setcell(&g, 0, 1, "7");
  setcell(&g, 0, 2, "11");
  setcell(&g, 0, 3, "+@SUM(A1...A3)");
  assert(g.cells[0][3].val == 23.0f);

  setcell(&g, 0, 0, "5");
  setcell(&g, 0, 1, "+A1+5");
  setcell(&g, 0, 2, "+A2+A1");
  assert(g.cells[0][3].val == 5.0f + 10.0f + 15.0f);

  setcell(&g, 0, 0, "7");
  assert(g.cells[0][3].val == 7.0f + 12.0f + 19.0f);
}

void test_ref(void) {
  int c, r;
  assert(ref("A1", &c, &r) == 2 && c == 0 && r == 0);
  assert(ref("Z50", &c, &r) == 3 && c == 25 && r == 49);
  assert(ref("AA10", &c, &r) == 4 && c == 26 && r == 9);
  assert(ref("AZ99", &c, &r) == 4 && c == 51 && r == 98);
  assert(ref("BA1", &c, &r) == 3 && c == 52 && r == 0);
}

void test_col(void) {
  assert(strcmp(col(0), "A") == 0);
  assert(strcmp(col(25), "Z") == 0);
  assert(strcmp(col(26), "AA") == 0);
  assert(strcmp(col(51), "AZ") == 0);
  assert(strcmp(col(52), "BA") == 0);
}

void test_csv_load(void) {
  struct grid g = {0};

  // basic CSV: numbers, labels, formulas
  {
    FILE* f = fopen("/tmp/test_kalk_basic.csv", "w");
    fprintf(f, "10,20,30\n");
    fprintf(f, "hello,world,+A1+B1\n");
    fclose(f);
    assert(csvload(&g, "/tmp/test_kalk_basic.csv") == 0);
    assert(g.cells[0][0].type == NUM && g.cells[0][0].val == 10.0f);
    assert(g.cells[1][0].type == NUM && g.cells[1][0].val == 20.0f);
    assert(g.cells[2][0].type == NUM && g.cells[2][0].val == 30.0f);
    assert(g.cells[0][1].type == LABEL);
    assert(strcmp(g.cells[0][1].text, "hello") == 0);
    assert(g.cells[2][1].type == FORMULA);
    assert(g.cells[2][1].val == 30.0f);  // A1+B1 = 10+20
  }

  // quoted fields: embedded commas and quotes
  {
    memset(&g, 0, sizeof(g));
    FILE* f = fopen("/tmp/test_kalk_quoted.csv", "w");
    fprintf(f, "\"has,comma\",\"has\"\"quote\",plain\n");
    fclose(f);
    assert(csvload(&g, "/tmp/test_kalk_quoted.csv") == 0);
    assert(strcmp(g.cells[0][0].text, "has,comma") == 0);
    assert(strcmp(g.cells[1][0].text, "has\"quote") == 0);
    assert(strcmp(g.cells[2][0].text, "plain") == 0);
  }

  // CRLF line endings (Windows/Excel)
  {
    memset(&g, 0, sizeof(g));
    FILE* f = fopen("/tmp/test_kalk_crlf.csv", "wb");
    fprintf(f, "1,2\r\n3,4\r\n");
    fclose(f);
    assert(csvload(&g, "/tmp/test_kalk_crlf.csv") == 0);
    assert(g.cells[0][0].val == 1.0f);
    assert(g.cells[1][0].val == 2.0f);
    assert(g.cells[0][1].val == 3.0f);
    assert(g.cells[1][1].val == 4.0f);
  }

  // empty cells
  {
    memset(&g, 0, sizeof(g));
    FILE* f = fopen("/tmp/test_kalk_empty.csv", "w");
    fprintf(f, "1,,3\n,5,\n");
    fclose(f);
    assert(csvload(&g, "/tmp/test_kalk_empty.csv") == 0);
    assert(g.cells[0][0].val == 1.0f);
    assert(g.cells[1][0].type == EMPTY);
    assert(g.cells[2][0].val == 3.0f);
    assert(g.cells[0][1].type == EMPTY);
    assert(g.cells[1][1].val == 5.0f);
  }

  // non-existent file
  assert(csvload(&g, "/tmp/test_kalk_nonexistent_42.csv") == -1);
}

void test_csv_save(void) {
  struct grid g = {0};

  setcell(&g, 0, 0, "100");
  setcell(&g, 1, 0, "200");
  setcell(&g, 2, 0, "+A1+B1");
  setcell(&g, 0, 1, "hello");
  setcell(&g, 1, 1, "has,comma");

  assert(csvsave(&g, "/tmp/test_kalk_save.csv") == 0);

  // read back and verify
  memset(&g, 0, sizeof(g));
  assert(csvload(&g, "/tmp/test_kalk_save.csv") == 0);
  assert(g.cells[0][0].val == 100.0f);
  assert(g.cells[1][0].val == 200.0f);
  assert(g.cells[2][0].type == FORMULA);
  assert(g.cells[2][0].val == 300.0f);
  assert(strcmp(g.cells[0][1].text, "hello") == 0);
  assert(strcmp(g.cells[1][1].text, "has,comma") == 0);
}

void test_csv_roundtrip(void) {
  struct grid g = {0};

  // field with embedded quote
  setcell(&g, 0, 0, "say \"hello\"");
  setcell(&g, 1, 0, "normal");
  setcell(&g, 0, 1, "42.5");

  assert(csvsave(&g, "/tmp/test_kalk_rt.csv") == 0);
  memset(&g, 0, sizeof(g));
  assert(csvload(&g, "/tmp/test_kalk_rt.csv") == 0);
  assert(strcmp(g.cells[0][0].text, "say \"hello\"") == 0);
  assert(strcmp(g.cells[1][0].text, "normal") == 0);
  assert(g.cells[0][1].val == 42.5f);
}

void test_swap(void) {
  struct grid g = {0};

  // Setup: A1=10 A2=20 A3=30, B1=100 B2=200 B3=300
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 0, 2, "30");
  setcell(&g, 1, 0, "100");
  setcell(&g, 1, 1, "200");
  setcell(&g, 1, 2, "300");

  // Swap rows 0 and 1: A1<->A2, B1<->B2
  swaprow(&g, 0, 1);
  assert(g.cells[0][0].val == 20.0f);
  assert(g.cells[0][1].val == 10.0f);
  assert(g.cells[1][0].val == 200.0f);
  assert(g.cells[1][1].val == 100.0f);
  assert(g.cells[0][2].val == 30.0f);  // row 2 unchanged

  // Swap back
  swaprow(&g, 0, 1);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[0][1].val == 20.0f);

  // Swap columns 0 and 1: A<->B
  swapcol(&g, 0, 1);
  assert(g.cells[0][0].val == 100.0f);
  assert(g.cells[1][0].val == 10.0f);
  assert(g.cells[0][1].val == 200.0f);
  assert(g.cells[1][1].val == 20.0f);

  // Swap back
  swapcol(&g, 0, 1);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[1][0].val == 100.0f);
}

void test_fixrefs(void) {
  struct grid g = {0};

  // Setup: A1=10, A2=20, formula in B1 referencing A1 and A2
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 1, 0, "+A1+A2");
  assert(g.cells[1][0].val == 30.0f);

  // Swap rows 0 and 1: data moves, refs updated
  // A1(10)->A2, A2(20)->A1, B1(+A1+A2)->B2 with refs swapped to +A2+A1
  swaprow(&g, 0, 1);
  recalc(&g);
  // Formula moved to B2, should still reference the same data (now A2+A1 = 10+20)
  assert(g.cells[1][1].val == 30.0f);
  assert(strcmp(g.cells[1][1].text, "+A2+A1") == 0);
  // Data swapped
  assert(g.cells[0][0].val == 20.0f);  // was A2
  assert(g.cells[0][1].val == 10.0f);  // was A1

  // Swap back: everything restored
  swaprow(&g, 0, 1);
  recalc(&g);
  assert(g.cells[1][0].val == 30.0f);
  assert(strcmp(g.cells[1][0].text, "+A1+A2") == 0);

  // Column swap: A1=10, B1=100, C1=+A1+B1
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 1, 0, "100");
  setcell(&g, 2, 0, "+A1+B1");
  assert(g.cells[2][0].val == 110.0f);

  // Swap cols A and B: data moves, refs updated
  swapcol(&g, 0, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 100.0f);  // was B
  assert(g.cells[1][0].val == 10.0f);   // was A
  assert(strcmp(g.cells[2][0].text, "+B1+A1") == 0);
  assert(g.cells[2][0].val == 110.0f);  // still 10+100

  // Swap back
  swapcol(&g, 0, 1);
  recalc(&g);
  assert(strcmp(g.cells[2][0].text, "+A1+B1") == 0);
  assert(g.cells[2][0].val == 110.0f);

  // Refs to uninvolved rows/cols stay unchanged
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  setcell(&g, 0, 2, "3");
  setcell(&g, 1, 0, "+A3");  // references row 2, not involved in swap of 0,1
  assert(g.cells[1][0].val == 3.0f);
  swaprow(&g, 0, 1);
  recalc(&g);
  // formula moved to B2, A3 ref unchanged since row 2 not swapped
  assert(strcmp(g.cells[1][1].text, "+A3") == 0);
  assert(g.cells[1][1].val == 3.0f);

  // SUM range refs get updated
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "5");
  setcell(&g, 0, 1, "10");
  setcell(&g, 0, 2, "15");
  setcell(&g, 1, 0, "+@SUM(A1...A3)");
  assert(g.cells[1][0].val == 30.0f);
  swaprow(&g, 0, 1);
  recalc(&g);
  // formula moved to B2, range refs swapped: A1->A2, A3 stays
  assert(strcmp(g.cells[1][1].text, "+@SUM(A2...A3)") == 0);
  assert(g.cells[1][1].val == 20.0f);  // A2(5)+A3(15), range shrank

  // Multiple swaps (move row 0 to row 2 via two swaps)
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  setcell(&g, 0, 2, "3");
  setcell(&g, 1, 0, "+A1");
  swaprow(&g, 0, 1);  // row0<->row1
  swaprow(&g, 1, 2);  // row1<->row2 (original row0 now at row2)
  recalc(&g);
  assert(g.cells[0][2].val == 1.0f);  // original A1 data
  assert(g.cells[1][2].val == 1.0f);  // formula followed the data
  assert(strcmp(g.cells[1][2].text, "+A3") == 0);

  // Formula not in swapped rows still gets refs updated
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 0, 2, "+A1+A2");  // in row 2, refs rows 0 and 1
  swaprow(&g, 0, 1);
  recalc(&g);
  assert(strcmp(g.cells[0][2].text, "+A2+A1") == 0);
  assert(g.cells[0][2].val == 30.0f);  // still 10+20
}

void test_insert_delete(void) {
  struct grid g = {0};

  // --- Insert row: data shifts down, refs adjust ---
  setcell(&g, 0, 0, "10");   // A1=10
  setcell(&g, 0, 1, "20");   // A2=20
  setcell(&g, 0, 2, "30");   // A3=30
  setcell(&g, 1, 0, "+A2");  // B1=+A2 (=20)
  assert(g.cells[1][0].val == 20.0f);

  // Insert row at 1 (between A1 and A2): A2->A3, A3->A4
  insertrow(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);   // A1 unchanged
  assert(g.cells[0][1].type == EMPTY);  // A2 is new blank row
  assert(g.cells[0][2].val == 20.0f);   // old A2 -> A3
  assert(g.cells[0][3].val == 30.0f);   // old A3 -> A4
  // B1 formula should now reference A3 (was A2, shifted +1)
  assert(strcmp(g.cells[1][0].text, "+A3") == 0);
  assert(g.cells[1][0].val == 20.0f);

  // --- Delete that inserted row: refs shrink back ---
  deleterow(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[0][1].val == 20.0f);
  assert(g.cells[0][2].val == 30.0f);
  assert(strcmp(g.cells[1][0].text, "+A2") == 0);
  assert(g.cells[1][0].val == 20.0f);

  // --- Insert column: data shifts right, refs adjust ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");      // A1=10
  setcell(&g, 1, 0, "20");      // B1=20
  setcell(&g, 2, 0, "+A1+B1");  // C1=+A1+B1 (=30)
  assert(g.cells[2][0].val == 30.0f);

  // Insert column at 1 (between A and B): B->C, C->D
  insertcol(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);   // A1 unchanged
  assert(g.cells[1][0].type == EMPTY);  // B1 is new blank col
  assert(g.cells[2][0].val == 20.0f);   // old B1 -> C1
  // Formula moved to D1, refs A1 unchanged, B1->C1
  assert(strcmp(g.cells[3][0].text, "+A1+C1") == 0);
  assert(g.cells[3][0].val == 30.0f);

  // --- Delete that inserted column ---
  deletecol(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[1][0].val == 20.0f);
  assert(strcmp(g.cells[2][0].text, "+A1+B1") == 0);
  assert(g.cells[2][0].val == 30.0f);

  // --- Delete row that is referenced: refs to later rows shift ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");   // A1=10
  setcell(&g, 0, 1, "20");   // A2=20
  setcell(&g, 0, 2, "30");   // A3=30
  setcell(&g, 1, 0, "+A3");  // B1=+A3 (=30)
  assert(g.cells[1][0].val == 30.0f);

  // Delete row 1 (A2): A3->A2
  deleterow(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[0][1].val == 30.0f);              // old A3 is now A2
  assert(strcmp(g.cells[1][0].text, "+A2") == 0);  // ref shifted
  assert(g.cells[1][0].val == 30.0f);

  // --- Insert at row 0: all refs shift ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "5");
  setcell(&g, 0, 1, "10");
  setcell(&g, 1, 1, "+A1");  // B2=+A1 (=5)
  assert(g.cells[1][1].val == 5.0f);

  insertrow(&g, 0);
  recalc(&g);
  assert(g.cells[0][0].type == EMPTY);  // new blank row
  assert(g.cells[0][1].val == 5.0f);    // old A1 -> A2
  assert(g.cells[0][2].val == 10.0f);   // old A2 -> A3
  // B2's formula was +A1, shifted to +A2 and moved to B3
  assert(strcmp(g.cells[1][2].text, "+A2") == 0);
  assert(g.cells[1][2].val == 5.0f);

  // --- SUM range adjusts on insert ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  setcell(&g, 0, 2, "3");
  setcell(&g, 1, 0, "+@SUM(A1...A3)");
  assert(g.cells[1][0].val == 6.0f);

  // Insert row at 1: range A1...A3 becomes A1...A4
  insertrow(&g, 1);
  recalc(&g);
  assert(strcmp(g.cells[1][0].text, "+@SUM(A1...A4)") == 0);
  assert(g.cells[1][0].val == 6.0f);  // new blank row adds 0

  // --- Delete column that is referenced ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");   // A1
  setcell(&g, 1, 0, "20");   // B1
  setcell(&g, 2, 0, "30");   // C1
  setcell(&g, 3, 0, "+C1");  // D1=+C1 (=30)
  assert(g.cells[3][0].val == 30.0f);

  // Delete col B (index 1): C->B, D->C, ref C1->B1
  deletecol(&g, 1);
  recalc(&g);
  assert(g.cells[0][0].val == 10.0f);
  assert(g.cells[1][0].val == 30.0f);              // old C1 is now B1
  assert(strcmp(g.cells[2][0].text, "+B1") == 0);  // ref shifted
  assert(g.cells[2][0].val == 30.0f);

  // --- Multiple inserts ---
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "42");
  setcell(&g, 1, 0, "+A1");
  insertrow(&g, 0);
  insertrow(&g, 0);
  recalc(&g);
  // A1 was shifted twice to A3
  assert(g.cells[0][2].val == 42.0f);
  assert(strcmp(g.cells[1][2].text, "+A3") == 0);
  assert(g.cells[1][2].val == 42.0f);
}

void test_replicate(void) {
  struct grid g = {0};

  // Replicate a number: plain copy
  setcell(&g, 0, 0, "42");
  replicatecell(&g, 0, 0, 1, 0);
  recalc(&g);
  assert(g.cells[1][0].type == NUM);
  assert(g.cells[1][0].val == 42.0f);

  // Replicate a label: plain copy
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "hello");
  replicatecell(&g, 0, 0, 1, 0);
  assert(g.cells[1][0].type == LABEL);
  assert(strcmp(g.cells[1][0].text, "hello") == 0);

  // Replicate formula down: refs shift by row delta
  // A1=10, A2=20, B1=+A1 → replicate B1 to B2 → B2=+A2
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 1, 0, "+A1");
  replicatecell(&g, 1, 0, 1, 1);  // B1 -> B2
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+A2") == 0);
  assert(g.cells[1][1].val == 20.0f);

  // Replicate formula right: refs shift by col delta
  // A1=10, B1=20, A2=+A1 → replicate A2 to B2 → B2=+B1
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 1, 0, "20");
  setcell(&g, 0, 1, "+A1");
  replicatecell(&g, 0, 1, 1, 1);  // A2 -> B2
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+B1") == 0);
  assert(g.cells[1][1].val == 20.0f);

  // Absolute ref: $A$1 stays fixed
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 1, 0, "+$A$1");
  replicatecell(&g, 1, 0, 1, 1);  // B1 -> B2
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+$A$1") == 0);
  assert(g.cells[1][1].val == 10.0f);

  // Mixed: $A1 → col fixed, row shifts
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 1, 0, "+$A1");
  replicatecell(&g, 1, 0, 1, 1);  // B1 -> B2
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+$A2") == 0);
  assert(g.cells[1][1].val == 20.0f);

  // Mixed: A$1 → row fixed, col shifts
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 1, 0, "20");
  setcell(&g, 0, 1, "+A$1");
  replicatecell(&g, 0, 1, 1, 1);  // A2 -> B2
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+B$1") == 0);
  assert(g.cells[1][1].val == 20.0f);

  // Replicate to range: B1=+A1, replicate to B2...B4
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  setcell(&g, 0, 2, "3");
  setcell(&g, 0, 3, "4");
  setcell(&g, 1, 0, "+A1");
  for (int r = 1; r <= 3; r++) replicatecell(&g, 1, 0, 1, r);
  recalc(&g);
  assert(strcmp(g.cells[1][1].text, "+A2") == 0);
  assert(g.cells[1][1].val == 2.0f);
  assert(strcmp(g.cells[1][2].text, "+A3") == 0);
  assert(g.cells[1][2].val == 3.0f);
  assert(strcmp(g.cells[1][3].text, "+A4") == 0);
  assert(g.cells[1][3].val == 4.0f);

  // SUM range shifts: B1=+@SUM(A1...A3), replicate to C1 → +@SUM(B1...B3)
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  setcell(&g, 0, 2, "3");
  setcell(&g, 1, 0, "+@SUM(A1...A3)");
  assert(g.cells[1][0].val == 6.0f);
  replicatecell(&g, 1, 0, 2, 0);  // B1 -> C1
  recalc(&g);
  assert(strcmp(g.cells[2][0].text, "+@SUM(B1...B3)") == 0);

  // Replicate empty cell clears target
  memset(&g, 0, sizeof(g));
  setcell(&g, 1, 0, "999");
  replicatecell(&g, 0, 0, 1, 0);  // A1 (empty) -> B1
  assert(g.cells[1][0].type == EMPTY);

  // $ refs survive in formula evaluation
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "7");
  setcell(&g, 1, 0, "+$A$1*2");
  recalc(&g);
  assert(g.cells[1][0].val == 14.0f);

  // Range-to-range: replicate A1...A3 to B1...B3
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "10");
  setcell(&g, 0, 1, "20");
  setcell(&g, 0, 2, "+A1+A2");
  assert(g.cells[0][2].val == 30.0f);
  // Simulate range replication: source A1...A3 -> target B1...B3
  {
    int sc1 = 0, sr1 = 0, sc2 = 0, sr2 = 2;
    int tc1 = 1, tr1 = 0, tc2 = 1, tr2 = 2;
    int sw = sc2 - sc1 + 1, sh = sr2 - sr1 + 1;
    for (int r = tr1; r <= tr2; r++)
      for (int c = tc1; c <= tc2; c++)
        replicatecell(&g, sc1 + (c - tc1) % sw, sr1 + (r - tr1) % sh, c, r);
    recalc(&g);
  }
  assert(g.cells[1][0].val == 10.0f);
  assert(g.cells[1][1].val == 20.0f);
  // B3 = +B1+B2 (shifted from +A1+A2)
  assert(strcmp(g.cells[1][2].text, "+B1+B2") == 0);
  assert(g.cells[1][2].val == 30.0f);

  // Range-to-range tiling: replicate A1...A2 to B1...B4 (tiles 2 rows into 4)
  memset(&g, 0, sizeof(g));
  setcell(&g, 0, 0, "1");
  setcell(&g, 0, 1, "2");
  {
    int sc1 = 0, sr1 = 0, sc2 = 0, sr2 = 1;
    int tc1 = 1, tr1 = 0, tc2 = 1, tr2 = 3;
    int sw = sc2 - sc1 + 1, sh = sr2 - sr1 + 1;
    for (int r = tr1; r <= tr2; r++)
      for (int c = tc1; c <= tc2; c++)
        replicatecell(&g, sc1 + (c - tc1) % sw, sr1 + (r - tr1) % sh, c, r);
    recalc(&g);
  }
  assert(g.cells[1][0].val == 1.0f);
  assert(g.cells[1][1].val == 2.0f);
  assert(g.cells[1][2].val == 1.0f);  // tiled
  assert(g.cells[1][3].val == 2.0f);  // tiled
}

int main(void) {
  test_expr();
  test_recalc();
  test_ref();
  test_col();
  test_csv_load();
  test_csv_save();
  test_csv_roundtrip();
  test_swap();
  test_fixrefs();
  test_insert_delete();
  test_replicate();
  return 0;
}
#endif
