all: kalk test

kalk:
	$(CC) -std=c99 -pedantic kalk.c -lcurses -o kalk

test:
	$(CC) -std=c99 -pedantic -DTEST=1 kalk.c -lcurses -o testkalk
	./testkalk

clean:
	rm -f kalk testkalk

.PHONY: all kalk test clean

