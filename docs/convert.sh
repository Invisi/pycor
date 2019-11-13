#!/usr/bin/env bash
VAR="--variable lang=de --variable papersize=a4 --variable documentclass=scrartcl"

pandoc $VAR -s Änderungen.md -o Änderungen.pdf

pandoc $VAR --toc -s Anleitung.md -o Anleitung.pdf