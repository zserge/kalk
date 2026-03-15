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
  int fmt;           // 0=general 'I'=integer '$'=dollar 'L'=left 'R'=right
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

int ref(const char* s, int* col, int* row) {
  char* end;
  const char* p = s;
  if (!isalpha(*p)) return 0;
  *col = toupper(*p++) - 'A' + 1;
  if (isalpha(*p)) *col = *col * 26 + (toupper(*p++) - 'A' + 1);
  int n = strtol(p, &end, 10);
  if (n <= 0 || end == p) return 0;
  *row = n - 1, *col = *col - 1;
  return (int)(end - s);
}

static float cellval(struct parser* p) {
  int c, r, n = ref(p->p, &c, &r);
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

static void draw(struct grid* g, const char* mode, const char* buf) {
  erase();

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
  for (int c = 0; c < vcols() && g->vc + c < NCOL; c++)
    mvprintw(2, GW + c * CW, "%*s", CW, col(g->vc + c));
  attroff(A_BOLD | A_REVERSE);

  // Grid cells
  for (int r = 0; r < vrows() && g->vr + r < NROW; r++) {
    int row = g->vr + r, y = 3 + r;
    move(y, 0);
    clrtoeol();
    attron(A_REVERSE);
    mvprintw(y, 0, "%*d ", GW - 1, row + 1);
    attroff(A_REVERSE);

    for (int c = 0; c < vcols() && g->vc + c < NCOL; c++) {
      int col = g->vc + c;
      struct cell* cl = cell(g, col, row);
      char fb[64];

      if (!cl || cl->type == EMPTY) {
        memset(fb, ' ', CW);
        fb[CW] = '\0';
      } else if (cl->type == LABEL)
        snprintf(fb, CW + 1, "%-*.*s", CW, CW, cl->text[0] == '"' ? cl->text + 1 : cl->text);
      else if (isnan(cl->val)) {
        snprintf(fb, CW + 1, "%*s", CW, "ERROR");
      } else {
        char t[MAXIN] = {0}, fmt = cl->fmt;
        if (!fmt || fmt == 'D')
          fmt = g->fmt;  // use default format
                         // L R I G D $ % *
        if (fmt == '$') {
          snprintf(t, sizeof(t), "%.2f", cl->val);
        } else if (fmt == '%') {
          snprintf(t, sizeof(t), "%.2f%%", cl->val * 100);
        } else if (fmt == '*') {
          for (int i = 0; i < CW && i < cl->val; i++) t[i] = '*';
          fmt = 'L';
        } else if (fmt == 'I' || (cl->val == (long)cl->val && fabs(cl->val) < 1e9)) {
          snprintf(t, sizeof(t), "%ld", (long)cl->val);
        } else {
          snprintf(t, sizeof(t), "%g", cl->val);
        }
        snprintf(fb, CW + 1, fmt == 'L' ? "%-*s" : "%*s", CW, t);
      }

      int is_cur = (col == g->cc && row == g->cr);
      if (is_cur) attron(A_REVERSE);
      mvprintw(y, GW + c * CW, "%s", fb);
      if (is_cur) attroff(A_REVERSE);
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
//  /M                   Move cell (cut/paste)
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
    if (ch == 'R') {
      for (int c = 0; c < NCOL; c++)
        for (int r = g->cr; r < NROW - 1; r++) g->cells[c][r] = g->cells[c][r + 1];
      for (int c = 0; c < NCOL; c++) g->cells[c][NROW - 1] = (struct cell){0};
      g->dirty = 1;
    } else if (ch == 'C') {
      for (int r = 0; r < NROW; r++)
        for (int c = g->cc; c < NCOL - 1; c++) g->cells[c][r] = g->cells[c + 1][r];
      for (int r = 0; r < NROW; r++) g->cells[NCOL - 1][r] = (struct cell){0};
      g->dirty = 1;
    }
  } else if (ch == 'I') {
    mvprintw(1, 0, "Insert (R)ow or (C)olumn?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'R') {
      for (int c = 0; c < NCOL; c++)
        for (int r = NROW - 1; r > g->cr; r--) g->cells[c][r] = g->cells[c][r - 1];
      for (int c = 0; c < NCOL; c++) g->cells[c][g->cr] = (struct cell){0};
      g->dirty = 1;
    } else if (ch == 'C') {
      for (int r = 0; r < NROW; r++)
        for (int c = NCOL - 1; c > g->cc; c--) g->cells[c][r] = g->cells[c - 1][r];
      for (int r = 0; r < NROW; r++) g->cells[g->cc][r] = (struct cell){0};
      g->dirty = 1;
    }
  } else if (ch == 'F') {
    mvprintw(1, 0, "Format: L R I G D $ %% *"), clrtoeol();
    ch = toupper(getch());
    struct cell* cl = cell(g, g->cc, g->cr);
    if (strchr("LRIGD$%*", ch)) cl->fmt = ch;
  } else if (ch == 'G') {
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
    //
  } else if (ch == 'R') {
    //
  } else if (ch == 'T') {
    mvprintw(1, 0, "Lock (V)ertical, (H)orizontal, (B)oth, or (N)one?"), clrtoeol();
    ch = toupper(getch());
    if (ch == 'V') {
      g->tc = g->cc, g->tr = -1;
    } else if (ch == 'H') {
      g->tr = g->cr, g->tc = -1;
    } else if (ch == 'B') {
      g->tc = g->cc, g->tr = g->cr;
    } else if (ch == 'N') {
      g->tc = g->tr = -1;
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
    if (g->cc < g->vc) g->vc = g->cc;
    if (g->cc >= g->vc + vcols()) g->vc = g->cc - vcols() + 1;
    if (g->cr < g->vr) g->vr = g->cr;
    if (g->cr >= g->vr + vrows()) g->vr = g->cr - vrows() + 1;
    draw(g, "READY", "");
    int ch = getch();

    if (ch == (0x1f & 'c'))  // Ctrl+C: quit
      break;
    else if (ch == KEY_UP && g->cr > 0)
      g->cr--;
    else if (ch == KEY_DOWN && g->cr < NROW - 1)
      g->cr++;
    else if (ch == KEY_LEFT && g->cc > 0)
      g->cc--;
    else if (ch == KEY_RIGHT && g->cc < NCOL - 1)
      g->cc++;
    else if (ch == KEY_HOME)
      g->cc = g->cr = 0;
    else if (ch == 9 && g->cc < NCOL - 1)  // Tab: next cell
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
    } else {
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

int main(void) {
  test_expr();
  test_recalc();
  test_ref();
  test_col();
  test_csv_load();
  test_csv_save();
  test_csv_roundtrip();
  return 0;
}
#endif
