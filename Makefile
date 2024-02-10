FILES = \
    Addons.xcu \
    META-INF/manifest.xml \
    description.xml \
    pkg-description/pkg-description.en \
    registration/license.txt \
    main.py \
    icons/*.bmp

oulipo.oxt: $(FILES)
	zip -r - $^ > $@ 
	
.PHONY: clean
clean:
	rm -f oulipo.oxt

.PHONY: f
f: clean oulipo.oxt
