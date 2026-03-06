import os
import io
import csv
from datetime import timedelta, datetime

from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from flask_jwt_extended import (
    JWTManager, create_access_token, jwt_required, get_jwt_identity
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------------------
# App Configuration
# ---------------------------------------------------------------------------
app = Flask(__name__)

# Render Disk mount path (default: /var/data)
DATA_DIR = os.environ.get("DATA_DIR", "/var/data")
os.makedirs(DATA_DIR, exist_ok=True)

app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DATA_DIR}/assets.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["JWT_SECRET_KEY"] = os.environ.get("JWT_SECRET_KEY", "change-me-in-production")
app.config["JWT_ACCESS_TOKEN_EXPIRES"] = timedelta(hours=8)

db = SQLAlchemy(app)
jwt = JWTManager(app)

# CORS – allow frontend origin(s)
FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "*")
CORS(app, resources={r"/api/*": {"origins": FRONTEND_ORIGIN}}, supports_credentials=True)

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    display_name = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Asset(db.Model):
    __tablename__ = "assets"
    id = db.Column(db.Integer, primary_key=True)
    management_code = db.Column(db.String(100), unique=True, nullable=False, index=True)
    name = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(50), nullable=False)          # medical / electronic / other
    category_other = db.Column(db.String(200), nullable=True)
    location = db.Column(db.String(100), nullable=False)
    location_other = db.Column(db.String(200), nullable=True)
    department = db.Column(db.String(100), nullable=False)
    department_other = db.Column(db.String(200), nullable=True)
    purchase_from = db.Column(db.String(200), nullable=True)
    purchase_price = db.Column(db.Integer, nullable=True)
    purchase_date = db.Column(db.String(20), nullable=True)
    has_maintenance = db.Column(db.Boolean, default=False)
    maintenance_info = db.Column(db.Text, nullable=True)
    maintenance_link = db.Column(db.String(500), nullable=True)
    depreciation_period_months = db.Column(db.Integer, nullable=True)
    lease_period_months = db.Column(db.Integer, nullable=True)
    manager = db.Column(db.String(100), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "management_code": self.management_code,
            "name": self.name,
            "category": self.category,
            "category_other": self.category_other,
            "location": self.location,
            "location_other": self.location_other,
            "department": self.department,
            "department_other": self.department_other,
            "purchase_from": self.purchase_from,
            "purchase_price": self.purchase_price,
            "purchase_date": self.purchase_date,
            "has_maintenance": self.has_maintenance,
            "maintenance_info": self.maintenance_info,
            "maintenance_link": self.maintenance_link,
            "depreciation_period_months": self.depreciation_period_months,
            "lease_period_months": self.lease_period_months,
            "manager": self.manager,
            "notes": self.notes,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
        }


class Department(db.Model):
    """設置部室マスタ – 各院で追加・削除可能"""
    __tablename__ = "departments"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    is_default = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "is_default": self.is_default,
        }


# ---------------------------------------------------------------------------
# Seed Data
# ---------------------------------------------------------------------------

DEFAULT_DEPARTMENTS = [
    "小児科1診", "小児科2診", "耳鼻科1診", "耳鼻科2診",
    "皮膚科診察室", "バックヤード", "受付", "その他",
]


def seed_data():
    """Create initial user and department records if they don't exist."""
    if User.query.count() == 0:
        admin = User(username="admin", display_name="管理者")
        admin.set_password("admin")
        db.session.add(admin)

    if Department.query.count() == 0:
        for name in DEFAULT_DEPARTMENTS:
            db.session.add(Department(name=name, is_default=True))

    db.session.commit()


# ---------------------------------------------------------------------------
# Auth Endpoints
# ---------------------------------------------------------------------------

@app.route("/api/auth/login", methods=["POST"])
def login():
    data = request.get_json(silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")

    user = User.query.filter_by(username=username).first()
    if not user or not user.check_password(password):
        return jsonify({"error": "ユーザー名またはパスワードが正しくありません"}), 401

    token = create_access_token(identity=str(user.id))
    return jsonify({
        "access_token": token,
        "user": {"id": user.id, "username": user.username, "display_name": user.display_name},
    })


@app.route("/api/auth/me", methods=["GET"])
@jwt_required()
def me():
    user = User.query.get(int(get_jwt_identity()))
    if not user:
        return jsonify({"error": "ユーザーが見つかりません"}), 404
    return jsonify({"id": user.id, "username": user.username, "display_name": user.display_name})


# ---------------------------------------------------------------------------
# Asset CRUD
# ---------------------------------------------------------------------------

@app.route("/api/assets", methods=["GET"])
@jwt_required()
def list_assets():
    q = Asset.query

    # Filtering
    category = request.args.get("category")
    location = request.args.get("location")
    search = request.args.get("search")

    if category:
        q = q.filter(Asset.category == category)
    if location:
        q = q.filter(Asset.location == location)
    if search:
        like = f"%{search}%"
        q = q.filter(
            db.or_(
                Asset.name.ilike(like),
                Asset.management_code.ilike(like),
                Asset.manager.ilike(like),
                Asset.purchase_from.ilike(like),
            )
        )

    # Sorting
    sort = request.args.get("sort", "updated_at")
    order = request.args.get("order", "desc")
    sort_col = getattr(Asset, sort, Asset.updated_at)
    q = q.order_by(sort_col.desc() if order == "desc" else sort_col.asc())

    # Pagination
    page = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 50, type=int)
    pagination = q.paginate(page=page, per_page=per_page, error_out=False)

    return jsonify({
        "items": [a.to_dict() for a in pagination.items],
        "total": pagination.total,
        "page": pagination.page,
        "per_page": pagination.per_page,
        "pages": pagination.pages,
    })


@app.route("/api/assets", methods=["POST"])
@jwt_required()
def create_asset():
    data = request.get_json(silent=True) or {}
    required = ["management_code", "name", "category", "location", "department"]
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({"error": f"必須項目が不足しています: {', '.join(missing)}"}), 400

    # Validate unique management_code
    if Asset.query.filter_by(management_code=data["management_code"]).first():
        return jsonify({"error": "この管理番号は既に登録されています"}), 409

    # Validate "その他" fields
    if data["location"] == "その他" and not data.get("location_other"):
        return jsonify({"error": "設置場所「その他」の場合は詳細を入力してください"}), 400
    if data["department"] == "その他" and not data.get("department_other"):
        return jsonify({"error": "設置部室「その他」の場合は詳細を入力してください"}), 400
    if data["category"] == "other" and not data.get("category_other"):
        return jsonify({"error": "種別「その他」の場合は詳細を入力してください"}), 400

    asset = Asset(
        management_code=data["management_code"],
        name=data["name"],
        category=data["category"],
        category_other=data.get("category_other"),
        location=data["location"],
        location_other=data.get("location_other"),
        department=data["department"],
        department_other=data.get("department_other"),
        purchase_from=data.get("purchase_from"),
        purchase_price=data.get("purchase_price"),
        purchase_date=data.get("purchase_date"),
        has_maintenance=data.get("has_maintenance", False),
        maintenance_info=data.get("maintenance_info"),
        maintenance_link=data.get("maintenance_link"),
        depreciation_period_months=data.get("depreciation_period_months"),
        lease_period_months=data.get("lease_period_months"),
        manager=data.get("manager"),
        notes=data.get("notes"),
    )
    db.session.add(asset)
    db.session.commit()
    return jsonify(asset.to_dict()), 201


@app.route("/api/assets/<int:asset_id>", methods=["GET"])
@jwt_required()
def get_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    return jsonify(asset.to_dict())


@app.route("/api/assets/<int:asset_id>", methods=["PUT"])
@jwt_required()
def update_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    data = request.get_json(silent=True) or {}

    # Check unique management_code if changed
    new_code = data.get("management_code", asset.management_code)
    if new_code != asset.management_code:
        if Asset.query.filter_by(management_code=new_code).first():
            return jsonify({"error": "この管理番号は既に登録されています"}), 409

    # Validate "その他"
    loc = data.get("location", asset.location)
    if loc == "その他" and not data.get("location_other", asset.location_other):
        return jsonify({"error": "設置場所「その他」の場合は詳細を入力してください"}), 400
    dept = data.get("department", asset.department)
    if dept == "その他" and not data.get("department_other", asset.department_other):
        return jsonify({"error": "設置部室「その他」の場合は詳細を入力してください"}), 400

    updatable = [
        "management_code", "name", "category", "category_other",
        "location", "location_other", "department", "department_other",
        "purchase_from", "purchase_price", "purchase_date",
        "has_maintenance", "maintenance_info", "maintenance_link",
        "depreciation_period_months", "lease_period_months",
        "manager", "notes",
    ]
    for field in updatable:
        if field in data:
            setattr(asset, field, data[field])

    db.session.commit()
    return jsonify(asset.to_dict())


@app.route("/api/assets/<int:asset_id>", methods=["DELETE"])
@jwt_required()
def delete_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    db.session.delete(asset)
    db.session.commit()
    return jsonify({"message": "削除しました"}), 200


# ---------------------------------------------------------------------------
# Bulk Download
# ---------------------------------------------------------------------------

CATEGORY_LABELS = {"medical": "医療機器", "electronic": "電子機器", "other": "その他"}

EXCEL_HEADERS = [
    ("管理番号", "management_code"),
    ("資産名", "name"),
    ("種別", "_category_display"),
    ("設置場所", "_location_display"),
    ("設置部室", "_department_display"),
    ("購入元", "purchase_from"),
    ("購入金額", "purchase_price"),
    ("購入日", "purchase_date"),
    ("保守有無", "_maintenance_display"),
    ("保守情報", "maintenance_info"),
    ("保守リンク", "maintenance_link"),
    ("減価償却期間(月)", "depreciation_period_months"),
    ("リース期間(月)", "lease_period_months"),
    ("管理者", "manager"),
    ("備考", "notes"),
    ("登録日", "created_at"),
    ("更新日", "updated_at"),
]


def _asset_row(asset):
    d = asset.to_dict()
    d["_category_display"] = CATEGORY_LABELS.get(d["category"], d.get("category_other") or d["category"])
    loc = d["location"]
    if loc == "その他" and d.get("location_other"):
        loc = f"その他({d['location_other']})"
    d["_location_display"] = loc
    dept = d["department"]
    if dept == "その他" and d.get("department_other"):
        dept = f"その他({d['department_other']})"
    d["_department_display"] = dept
    d["_maintenance_display"] = "有" if d["has_maintenance"] else "無"
    return d


@app.route("/api/assets/download", methods=["GET"])
@jwt_required()
def download_assets():
    fmt = request.args.get("format", "xlsx")
    assets = Asset.query.order_by(Asset.management_code).all()

    if fmt == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow([h[0] for h in EXCEL_HEADERS])
        for asset in assets:
            row = _asset_row(asset)
            writer.writerow([row.get(h[1], "") for h in EXCEL_HEADERS])
        output.seek(0)
        buf = io.BytesIO(output.getvalue().encode("utf-8-sig"))
        return send_file(buf, mimetype="text/csv", as_attachment=True,
                         download_name=f"assets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

    # Default: Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "資産一覧"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2B5797", end_color="2B5797", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    for col_idx, (header, _) in enumerate(EXCEL_HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for row_idx, asset in enumerate(assets, 2):
        row_data = _asset_row(asset)
        for col_idx, (_, key) in enumerate(EXCEL_HEADERS, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row_data.get(key, ""))
            cell.border = thin_border

    # Auto column width
    for col_idx, (header, _) in enumerate(EXCEL_HEADERS, 1):
        ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else "A"].width = max(len(header) * 2.5, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=f"assets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


# ---------------------------------------------------------------------------
# CSV Import (Bulk Upload)
# ---------------------------------------------------------------------------

IMPORT_HEADERS = [
    "管理番号", "資産名", "種別", "種別(その他)", "設置場所", "設置場所(その他)",
    "設置部室", "設置部室(その他)", "購入元", "購入金額", "購入日",
    "保守有無", "保守情報", "保守リンク", "減価償却期間(月)", "リース期間(月)",
    "管理者", "備考",
]

CATEGORY_REVERSE = {"医療機器": "medical", "電子機器": "electronic", "その他": "other"}


@app.route("/api/assets/upload-template", methods=["GET"])
@jwt_required()
def download_template():
    """Download empty CSV template for bulk import."""
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(IMPORT_HEADERS)
    # Write one example row
    writer.writerow([
        "MED-001", "超音波診断装置", "医療機器", "", "豊洲院", "",
        "小児科1診", "", "GEヘルスケア", "4800000", "2024-04-15",
        "有", "年間保守契約", "https://example.com", "72", "",
        "佐藤医師", "備考メモ",
    ])
    output.seek(0)
    buf = io.BytesIO(output.getvalue().encode("utf-8-sig"))
    return send_file(buf, mimetype="text/csv", as_attachment=True,
                     download_name="import_template.csv")


@app.route("/api/assets/upload", methods=["POST"])
@jwt_required()
def upload_assets():
    """Bulk import assets from CSV file."""
    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    file = request.files["file"]
    if not file.filename.endswith(".csv"):
        return jsonify({"error": "CSVファイルのみ対応しています"}), 400

    try:
        # Read CSV with BOM handling
        raw = file.read()
        # Try utf-8-sig first, then shift_jis
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("shift_jis")

        reader = csv.DictReader(io.StringIO(text))

        created = 0
        skipped = 0
        errors = []

        for i, row in enumerate(reader, start=2):
            code = (row.get("管理番号") or "").strip()
            name = (row.get("資産名") or "").strip()

            if not code or not name:
                skipped += 1
                continue

            # Skip if management_code already exists
            if Asset.query.filter_by(management_code=code).first():
                errors.append(f"行{i}: 管理番号「{code}」は既に登録済み")
                skipped += 1
                continue

            # Parse category
            cat_label = (row.get("種別") or "").strip()
            category = CATEGORY_REVERSE.get(cat_label, "other")
            category_other = (row.get("種別(その他)") or "").strip()
            if cat_label and cat_label not in CATEGORY_REVERSE:
                category = "other"
                category_other = cat_label

            # Parse location
            location = (row.get("設置場所") or "").strip()
            location_other = (row.get("設置場所(その他)") or "").strip()
            if not location:
                errors.append(f"行{i}: 設置場所が空です")
                skipped += 1
                continue

            # Parse department (default to "その他" if empty)
            department = (row.get("設置部室") or "").strip()
            department_other = (row.get("設置部室(その他)") or "").strip()
            if not department:
                department = "その他"

            # Parse numeric fields
            def parse_int(val):
                val = (val or "").strip().replace(",", "")
                if not val:
                    return None
                try:
                    return int(float(val))
                except (ValueError, TypeError):
                    return None

            # Parse maintenance flag
            maint_str = (row.get("保守有無") or "").strip()
            has_maintenance = maint_str in ("有", "あり", "○", "1", "true", "True", "YES", "yes")

            asset = Asset(
                management_code=code,
                name=name,
                category=category,
                category_other=category_other or None,
                location=location,
                location_other=location_other or None,
                department=department,
                department_other=department_other or None,
                purchase_from=(row.get("購入元") or "").strip() or None,
                purchase_price=parse_int(row.get("購入金額")),
                purchase_date=(row.get("購入日") or "").strip() or None,
                has_maintenance=has_maintenance,
                maintenance_info=(row.get("保守情報") or "").strip() or None,
                maintenance_link=(row.get("保守リンク") or "").strip() or None,
                depreciation_period_months=parse_int(row.get("減価償却期間(月)")),
                lease_period_months=parse_int(row.get("リース期間(月)")),
                manager=(row.get("管理者") or "").strip() or None,
                notes=(row.get("備考") or "").strip() or None,
            )
            db.session.add(asset)
            created += 1

        db.session.commit()

        return jsonify({
            "message": f"{created}件を登録しました",
            "created": created,
            "skipped": skipped,
            "errors": errors[:20],  # Return first 20 errors max
        }), 200

    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"CSVの処理中にエラーが発生しました: {str(e)}"}), 400


# ---------------------------------------------------------------------------
# Department CRUD
# ---------------------------------------------------------------------------

@app.route("/api/departments", methods=["GET"])
@jwt_required()
def list_departments():
    depts = Department.query.order_by(Department.id).all()
    return jsonify([d.to_dict() for d in depts])


@app.route("/api/departments", methods=["POST"])
@jwt_required()
def create_department():
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip()
    if not name:
        return jsonify({"error": "部室名を入力してください"}), 400
    if Department.query.filter_by(name=name).first():
        return jsonify({"error": "同名の部室が既に存在します"}), 409
    dept = Department(name=name, is_default=False)
    db.session.add(dept)
    db.session.commit()
    return jsonify(dept.to_dict()), 201


@app.route("/api/departments/<int:dept_id>", methods=["DELETE"])
@jwt_required()
def delete_department(dept_id):
    dept = Department.query.get_or_404(dept_id)
    db.session.delete(dept)
    db.session.commit()
    return jsonify({"message": "削除しました"}), 200


# ---------------------------------------------------------------------------
# Stats / Dashboard
# ---------------------------------------------------------------------------

@app.route("/api/stats", methods=["GET"])
@jwt_required()
def stats():
    total = Asset.query.count()
    by_category = db.session.query(
        Asset.category, db.func.count(Asset.id)
    ).group_by(Asset.category).all()
    by_location = db.session.query(
        Asset.location, db.func.count(Asset.id)
    ).group_by(Asset.location).all()
    maintenance_count = Asset.query.filter_by(has_maintenance=True).count()

    return jsonify({
        "total": total,
        "by_category": {CATEGORY_LABELS.get(c, c): n for c, n in by_category},
        "by_location": dict(by_location),
        "maintenance_count": maintenance_count,
    })


# ---------------------------------------------------------------------------
# Health Check
# ---------------------------------------------------------------------------

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "timestamp": datetime.utcnow().isoformat()})


# ---------------------------------------------------------------------------
# Frontend (serve index.html)
# ---------------------------------------------------------------------------

@app.route("/")
def serve_frontend():
    return send_file("index.html")

# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------

with app.app_context():
    db.create_all()
    seed_data()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
