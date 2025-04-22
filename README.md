```
# pdf_en_to_ja-main/
##  ├── in/            # 入力PDFを配置するフォルダ（なければ作成）
##  ├── out/           # 出力ファイルが保存されるフォルダ（なければ作成）
##  ├── python_pdf_en_to_ja_m.py
##  └── requirements.txt
```

```
pip install -r requirements.txt
```

``` requirements.txt
# PDFファイル処理用
PyMuPDF>=1.19.0

# 翻訳用
deep-translator>=1.9.0

# 自然言語処理用
nltk>=3.6.0

# Word文書作成用
python-docx>=0.8.11
```
