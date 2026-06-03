#!/usr/bin/env python3
"""公文正文排版脚本 — 按 GB/T 9704-2012 格式生成 .docx"""
import argparse,json,os,re,sys
from pathlib import Path
if sys.stdout.encoding!='utf-8':sys.stdout.reconfigure(encoding='utf-8');sys.stderr.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt,Cm,Mm,Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
