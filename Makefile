PREFIX ?= /usr/local
BINDIR ?= $(PREFIX)/bin
MANDIR ?= $(PREFIX)/share/man/man1

all: kalk test

kalk:
	$(CC) -std=c99 -pedantic kalk.c -lcurses -lm -o kalk

test:
	$(CC) -std=c99 -pedantic -DTEST=1 kalk.c -lcurses -lm -o testkalk
	./testkalk

install: kalk
	install -d $(DESTDIR)$(BINDIR)
	install -m 755 kalk $(DESTDIR)$(BINDIR)/kalk
	install -d $(DESTDIR)$(MANDIR)
	install -m 644 kalk.1 $(DESTDIR)$(MANDIR)/kalk.1

uninstall:
	rm -f $(DESTDIR)$(BINDIR)/kalk
	rm -f $(DESTDIR)$(MANDIR)/kalk.1

clean:
	rm -f kalk testkalk

.PHONY: all kalk test install uninstall clean

