CURDIR=$(shell pwd)
LOCLDIR=/usr/local/bin

install: sc

update: sc

sc: .FORCE
	cp -fp $(CURDIR)/sc.py $(LOCLDIR)/sc.py

clean:
	rm -f $(LOCLDIR)/sc.py

.FORCE:


