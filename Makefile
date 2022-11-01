all: maquette.tex smalatex.py
	python smalatex.py
	pdflatex maquette.tex