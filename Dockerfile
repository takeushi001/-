# ベースイメージとしてPythonを使用
FROM python:3.9-slim

# 必要なパッケージをインストール
RUN apt-get update && apt-get install -y gcc

# 作業ディレクトリを設定
WORKDIR /app

# 必要なパッケージをインストール
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# スクリプトをコンテナにコピー
COPY bukkaku_automation.py .

# コンテナ起動時に実行するコマンドを設定
CMD ["python", "bukkaku_automation.py"]