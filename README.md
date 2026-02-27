# 資産管理システム - バックエンド API

## 概要
Flask + SQLite で構築した REST API。Render の Web Service + Disk で永続化。

## ローカル開発

```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

API は `http://localhost:5000` で起動します。

## Render デプロイ手順

### 1. GitHubリポジトリを作成して push
```bash
git init && git add . && git commit -m "init"
git remote add origin https://github.com/<user>/asset-management-api.git
git push -u origin main
```

### 2. Render で Web Service を作成
1. https://render.com → **New → Web Service**
2. GitHub リポジトリを接続
3. 設定:
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`

### 3. Disk を追加 (永続化)
1. Web Service の **Disks** タブ → **Add Disk**
2. **Name**: `asset-data`
3. **Mount Path**: `/var/data`
4. **Size**: 1 GB

### 4. 環境変数を設定
| Key | Value |
|-----|-------|
| `JWT_SECRET_KEY` | ランダムな文字列 (Render の Generate で可) |
| `DATA_DIR` | `/var/data` |
| `FRONTEND_ORIGIN` | `https://<username>.github.io` |
| `PYTHON_VERSION` | `3.11.0` |

### 5. デプロイ確認
```bash
curl https://<your-service>.onrender.com/api/health
```

## API エンドポイント一覧

| Method | Path | 説明 |
|--------|------|------|
| POST | `/api/auth/login` | ログイン |
| GET | `/api/auth/me` | ログインユーザー情報 |
| GET | `/api/assets` | 資産一覧 (検索・フィルタ・ページネーション) |
| POST | `/api/assets` | 資産登録 |
| GET | `/api/assets/<id>` | 資産詳細 |
| PUT | `/api/assets/<id>` | 資産更新 |
| DELETE | `/api/assets/<id>` | 資産削除 |
| GET | `/api/assets/download?format=xlsx\|csv` | 一括ダウンロード |
| GET | `/api/departments` | 部室一覧 |
| POST | `/api/departments` | 部室追加 |
| DELETE | `/api/departments/<id>` | 部室削除 |
| GET | `/api/stats` | ダッシュボード統計 |
| GET | `/api/health` | ヘルスチェック |

## 初期データ
- ユーザー: `admin` / `admin`
- 設置部室: 小児科1診, 小児科2診, 耳鼻科1診, 耳鼻科2診, 皮膚科診察室, バックヤード, 受付, その他
