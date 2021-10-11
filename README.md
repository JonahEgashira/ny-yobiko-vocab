## How to use

### Needed
- python, pipenv
- Libreoffice

### Steps
- $ pipenv install
- $ pipenv shell

- create.py -> 2ページ分の単語テストを5部ずつ作成, 移動先のディレクトリを作成
- 例: $ python create.py five で5級のテスト作成
- convert.shでxlsxをpdfに変換
- move.py -> ./のxlsx, pdfファイルを作成されたディレクトリに移動
- 例: python move.py 14(page_begin) 196(page_end)
- singlepage.py -> 単語テストのまとめを作成
- 例: python singlepage.py four で4級のテスト作成
- random_quiz.py -> grade_page.csvで指定されたページ範囲の単語の中から20単語を選んでまとめテスト作成
- 例: python random_quiz.py pre-two で準2級のまとめテストを作成

