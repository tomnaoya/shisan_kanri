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

DATA_DIR = os.environ.get("DATA_DIR", "/var/data")
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except PermissionError:
    DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    os.makedirs(DATA_DIR, exist_ok=True)

app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DATA_DIR}/assets.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["JWT_SECRET_KEY"] = os.environ.get("JWT_SECRET_KEY", "change-me-in-production")
app.config["JWT_ACCESS_TOKEN_EXPIRES"] = timedelta(hours=8)

db = SQLAlchemy(app)
jwt = JWTManager(app)

CORS(app, resources={r"/api/*": {"origins": "*"}}, supports_credentials=False)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Category prefix mapping for auto-numbering
CATEGORY_PREFIX = {
    "medical":        "A1",
    "medical_supply": "B1",
    "electronic":     "C1",
    "intangible":     "N1",
    "other":          "Z1",
}

CATEGORY_LABELS = {
    "medical":        "医療機器",
    "medical_supply": "医療資材",
    "electronic":     "電子機器",
    "intangible":     "非実在資産",
    "other":          "その他",
}

CATEGORY_REVERSE = {v: k for k, v in CATEGORY_LABELS.items()}

# Maintenance options: "有", "無", "不明"
MAINTENANCE_VALUES = {"有", "無", "不明"}

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    display_name = db.Column(db.String(120), nullable=True)
    role = db.Column(db.String(20), nullable=False, default="user")  # admin / user
    location = db.Column(db.String(100), nullable=True)  # belonging location (null = all for admin)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "display_name": self.display_name,
            "role": self.role,
            "location": self.location,
            "created_at": self.created_at.isoformat() if self.created_at else None,
        }


class Asset(db.Model):
    __tablename__ = "assets"
    id = db.Column(db.Integer, primary_key=True)
    management_code = db.Column(db.String(100), unique=True, nullable=False, index=True)
    name = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(50), nullable=False)
    category_other = db.Column(db.String(200), nullable=True)
    location = db.Column(db.String(100), nullable=False)
    location_other = db.Column(db.String(200), nullable=True)
    department = db.Column(db.String(100), nullable=False)
    department_other = db.Column(db.String(200), nullable=True)
    purchase_from = db.Column(db.String(200), nullable=True)
    purchase_price = db.Column(db.Integer, nullable=True)
    purchase_date = db.Column(db.String(20), nullable=True)
    maintenance_status = db.Column(db.String(10), default="無")  # 有 / 無 / 不明
    maintenance_info = db.Column(db.Text, nullable=True)
    maintenance_link = db.Column(db.String(500), nullable=True)
    depreciation_period_months = db.Column(db.Integer, nullable=True)
    lease_period_months = db.Column(db.Integer, nullable=True)
    manager = db.Column(db.String(100), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    is_deleted = db.Column(db.Boolean, default=False, index=True)
    deleted_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Migration helper: old has_maintenance boolean
    has_maintenance = db.Column(db.Boolean, nullable=True)

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
            "maintenance_status": self.maintenance_status or "無",
            "maintenance_info": self.maintenance_info,
            "maintenance_link": self.maintenance_link,
            "depreciation_period_months": self.depreciation_period_months,
            "lease_period_months": self.lease_period_months,
            "manager": self.manager,
            "notes": self.notes,
            "is_deleted": self.is_deleted,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
        }


class Department(db.Model):
    __tablename__ = "departments"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    is_default = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {"id": self.id, "name": self.name, "is_default": self.is_default}


class Location(db.Model):
    __tablename__ = "locations"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    is_default = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {"id": self.id, "name": self.name, "is_default": self.is_default}


# ---------------------------------------------------------------------------
# Auto-numbering
# ---------------------------------------------------------------------------

def generate_management_code(category):
    """Generate next management_code for given category."""
    prefix = CATEGORY_PREFIX.get(category, "Z1")
    # Find max existing code with this prefix (including deleted)
    like_pattern = f"{prefix}-%"
    last = (
        Asset.query
        .filter(Asset.management_code.like(like_pattern))
        .order_by(Asset.management_code.desc())
        .first()
    )
    if last:
        try:
            num = int(last.management_code.split("-")[1]) + 1
        except (IndexError, ValueError):
            num = 1
    else:
        num = 1
    return f"{prefix}-{num:05d}"


# ---------------------------------------------------------------------------
# Seed Data
# ---------------------------------------------------------------------------

DEFAULT_DEPARTMENTS = [
    "小児科1診", "小児科2診", "耳鼻科1診", "耳鼻科2診",
    "皮膚科診察室", "バックヤード", "受付", "その他",
]

DEFAULT_LOCATIONS = [
    "豊洲院", "勝どき院", "田町芝浦院", "ガーデン院小児耳鼻",
    "ガーデン院皮膚", "柏院", "有明院", "有明ひふか院", "サポートチーム",
]


def seed_data():
    if User.query.count() == 0:
        admin = User(username="admin", display_name="管理者", role="admin", location=None)
        admin.set_password("admin")
        db.session.add(admin)

    if Department.query.count() == 0:
        for name in DEFAULT_DEPARTMENTS:
            db.session.add(Department(name=name, is_default=True))

    if Location.query.count() == 0:
        for name in DEFAULT_LOCATIONS:
            db.session.add(Location(name=name, is_default=True))

    db.session.commit()


def migrate_data():
    """Migrate existing data to new schema."""
    # Add role/location columns to existing admin users
    for user in User.query.all():
        if not user.role:
            user.role = "admin"

    # Migrate has_maintenance -> maintenance_status
    for asset in Asset.query.all():
        if asset.maintenance_status is None or asset.maintenance_status == "":
            if asset.has_maintenance is True:
                asset.maintenance_status = "有"
            elif asset.has_maintenance is False:
                asset.maintenance_status = "無"
            else:
                asset.maintenance_status = "不明"
        if asset.is_deleted is None:
            asset.is_deleted = False

    db.session.commit()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_current_user():
    uid = int(get_jwt_identity())
    return User.query.get(uid)


def check_location_permission(user, asset_location):
    """Check if user has permission for given location."""
    if user.role == "admin":
        return True
    if user.location == "サポートチーム":
        return True
    return user.location == asset_location


# ---------------------------------------------------------------------------
# Auth
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
    return jsonify({"access_token": token, "user": user.to_dict()})


@app.route("/api/auth/me", methods=["GET"])
@jwt_required()
def me():
    user = get_current_user()
    if not user:
        return jsonify({"error": "ユーザーが見つかりません"}), 404
    return jsonify(user.to_dict())


# ---------------------------------------------------------------------------
# Asset CRUD
# ---------------------------------------------------------------------------

@app.route("/api/assets", methods=["GET"])
@jwt_required()
def list_assets():
    user = get_current_user()
    q = Asset.query.filter_by(is_deleted=False)

    # Location permission
    if user.role != "admin" and user.location != "サポートチーム":
        q = q.filter(Asset.location == user.location)

    # Filters
    category = request.args.get("category")
    location = request.args.get("location")
    search = request.args.get("search")
    date_from = request.args.get("date_from")
    date_to = request.args.get("date_to")

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
    if date_from:
        q = q.filter(Asset.purchase_date >= date_from)
    if date_to:
        q = q.filter(Asset.purchase_date <= date_to)

    # Sort
    sort = request.args.get("sort", "updated_at")
    order = request.args.get("order", "desc")
    sort_col = getattr(Asset, sort, Asset.updated_at)
    q = q.order_by(sort_col.desc() if order == "desc" else sort_col.asc())

    # Paginate
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


@app.route("/api/assets/next-code/<category>", methods=["GET"])
@jwt_required()
def next_code(category):
    """Preview next management code for a category."""
    code = generate_management_code(category)
    return jsonify({"code": code})


@app.route("/api/assets", methods=["POST"])
@jwt_required()
def create_asset():
    user = get_current_user()
    data = request.get_json(silent=True) or {}

    required = ["name", "category", "location"]
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({"error": f"必須項目が不足しています: {', '.join(missing)}"}), 400

    loc = data["location"]
    if not check_location_permission(user, loc):
        return jsonify({"error": "この設置場所への登録権限がありません"}), 403

    # Validate
    if loc == "その他" and not data.get("location_other"):
        return jsonify({"error": "設置場所「その他」の場合は詳細を入力してください"}), 400
    if data.get("category") == "other" and not data.get("category_other"):
        return jsonify({"error": "種別「その他」の場合は詳細を入力してください"}), 400

    dept = data.get("department", "その他")
    if not dept:
        dept = "その他"

    # Auto-generate management code
    code = generate_management_code(data["category"])

    maint = data.get("maintenance_status", "無")
    if maint not in MAINTENANCE_VALUES:
        maint = "無"

    asset = Asset(
        management_code=code,
        name=data["name"],
        category=data["category"],
        category_other=data.get("category_other"),
        location=loc,
        location_other=data.get("location_other"),
        department=dept,
        department_other=data.get("department_other"),
        purchase_from=data.get("purchase_from"),
        purchase_price=data.get("purchase_price"),
        purchase_date=data.get("purchase_date"),
        maintenance_status=maint,
        maintenance_info=data.get("maintenance_info"),
        maintenance_link=data.get("maintenance_link"),
        depreciation_period_months=data.get("depreciation_period_months"),
        lease_period_months=data.get("lease_period_months"),
        manager=data.get("manager"),
        notes=data.get("notes"),
        is_deleted=False,
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
    user = get_current_user()
    asset = Asset.query.get_or_404(asset_id)

    if asset.is_deleted:
        return jsonify({"error": "削除済みの資産は編集できません"}), 400

    if not check_location_permission(user, asset.location):
        return jsonify({"error": "この資産の編集権限がありません"}), 403

    data = request.get_json(silent=True) or {}

    # If category changed, we do NOT change management_code (already assigned)
    # Validate location change permission
    new_loc = data.get("location", asset.location)
    if new_loc != asset.location and not check_location_permission(user, new_loc):
        return jsonify({"error": "移動先の設置場所への権限がありません"}), 403

    updatable = [
        "name", "category", "category_other",
        "location", "location_other", "department", "department_other",
        "purchase_from", "purchase_price", "purchase_date",
        "maintenance_status", "maintenance_info", "maintenance_link",
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
    """Soft delete – mark as deleted, never physically remove."""
    user = get_current_user()
    asset = Asset.query.get_or_404(asset_id)

    if not check_location_permission(user, asset.location):
        return jsonify({"error": "この資産の削除権限がありません"}), 403

    asset.is_deleted = True
    asset.deleted_at = datetime.utcnow()
    db.session.commit()
    return jsonify({"message": "削除しました"}), 200


# ---------------------------------------------------------------------------
# Bulk Download
# ---------------------------------------------------------------------------

EXCEL_HEADERS = [
    ("管理番号", "management_code"),
    ("資産名", "name"),
    ("種別", "_category_display"),
    ("設置場所", "_location_display"),
    ("設置部室", "_department_display"),
    ("購入元", "purchase_from"),
    ("購入金額", "purchase_price"),
    ("購入日", "purchase_date"),
    ("保守", "_maintenance_display"),
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
    d["_maintenance_display"] = d.get("maintenance_status", "無")
    return d


@app.route("/api/assets/download", methods=["GET"])
@jwt_required()
def download_assets():
    fmt = request.args.get("format", "xlsx")
    assets = Asset.query.filter_by(is_deleted=False).order_by(Asset.management_code).all()

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
    for col_idx, (header, _) in enumerate(EXCEL_HEADERS, 1):
        ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else "A"].width = max(len(header) * 2.5, 12)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=f"assets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


# ---------------------------------------------------------------------------
# CSV Import
# ---------------------------------------------------------------------------

IMPORT_HEADERS = [
    "管理番号", "資産名", "種別", "種別(その他)", "設置場所", "設置場所(その他)",
    "設置部室", "設置部室(その他)", "購入元", "購入金額", "購入日",
    "保守有無", "保守情報", "保守リンク", "減価償却期間(月)", "リース期間(月)",
    "管理者", "備考",
]


@app.route("/api/assets/upload-template", methods=["GET"])
@jwt_required()
def download_template():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(IMPORT_HEADERS)
    writer.writerow([
        "(自動採番)", "超音波診断装置", "医療機器", "", "豊洲院", "",
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
    user = get_current_user()

    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400
    file = request.files["file"]
    if not file.filename.endswith(".csv"):
        return jsonify({"error": "CSVファイルのみ対応しています"}), 400

    try:
        raw = file.read()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("shift_jis")

        reader = csv.DictReader(io.StringIO(text))
        created = 0
        skipped = 0
        errors = []

        for i, row in enumerate(reader, start=2):
            name = (row.get("資産名") or "").strip()
            if not name:
                skipped += 1
                continue

            cat_label = (row.get("種別") or "").strip()
            category = CATEGORY_REVERSE.get(cat_label, "other")
            category_other = (row.get("種別(その他)") or "").strip()
            if cat_label and cat_label not in CATEGORY_REVERSE:
                category = "other"
                category_other = cat_label

            location = (row.get("設置場所") or "").strip()
            location_other = (row.get("設置場所(その他)") or "").strip()
            if not location:
                errors.append(f"行{i}: 設置場所が空です")
                skipped += 1
                continue

            if not check_location_permission(user, location):
                errors.append(f"行{i}: {location}への登録権限がありません")
                skipped += 1
                continue

            department = (row.get("設置部室") or "").strip()
            department_other = (row.get("設置部室(その他)") or "").strip()
            if not department:
                department = "その他"

            def parse_int(val):
                val = (val or "").strip().replace(",", "")
                if not val:
                    return None
                try:
                    return int(float(val))
                except (ValueError, TypeError):
                    return None

            maint_str = (row.get("保守有無") or "").strip()
            if maint_str in ("有", "あり", "○", "1", "true", "True", "YES", "yes"):
                maint = "有"
            elif maint_str in ("不明",):
                maint = "不明"
            else:
                maint = "無"

            # Auto-generate code
            code = generate_management_code(category)

            # Put old management_code (from CSV) into notes
            old_code = (row.get("管理番号") or "").strip()
            notes_val = (row.get("備考") or "").strip()
            if old_code:
                if notes_val:
                    notes_val = f"旧管理番号: {old_code} / {notes_val}"
                else:
                    notes_val = f"旧管理番号: {old_code}"

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
                maintenance_status=maint,
                maintenance_info=(row.get("保守情報") or "").strip() or None,
                maintenance_link=(row.get("保守リンク") or "").strip() or None,
                depreciation_period_months=parse_int(row.get("減価償却期間(月)")),
                lease_period_months=parse_int(row.get("リース期間(月)")),
                manager=(row.get("管理者") or "").strip() or None,
                notes=notes_val or None,
                is_deleted=False,
            )
            db.session.add(asset)
            created += 1

        db.session.commit()
        return jsonify({
            "message": f"{created}件を登録しました",
            "created": created,
            "skipped": skipped,
            "errors": errors[:20],
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
# Location CRUD (admin only)
# ---------------------------------------------------------------------------

@app.route("/api/locations", methods=["GET"])
@jwt_required()
def list_locations():
    locs = Location.query.order_by(Location.id).all()
    return jsonify([l.to_dict() for l in locs])


@app.route("/api/locations", methods=["POST"])
@jwt_required()
def create_location():
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip()
    if not name:
        return jsonify({"error": "設置場所名を入力してください"}), 400
    if Location.query.filter_by(name=name).first():
        return jsonify({"error": "同名の設置場所が既に存在します"}), 409
    loc = Location(name=name, is_default=False)
    db.session.add(loc)
    db.session.commit()
    return jsonify(loc.to_dict()), 201


@app.route("/api/locations/<int:loc_id>", methods=["PUT"])
@jwt_required()
def update_location(loc_id):
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    loc = Location.query.get_or_404(loc_id)
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip()
    if name and name != loc.name:
        if Location.query.filter_by(name=name).first():
            return jsonify({"error": "同名の設置場所が既に存在します"}), 409
        loc.name = name
    db.session.commit()
    return jsonify(loc.to_dict())


@app.route("/api/locations/<int:loc_id>", methods=["DELETE"])
@jwt_required()
def delete_location(loc_id):
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    loc = Location.query.get_or_404(loc_id)
    db.session.delete(loc)
    db.session.commit()
    return jsonify({"message": "削除しました"}), 200


# ---------------------------------------------------------------------------
# User Management (admin only)
# ---------------------------------------------------------------------------

@app.route("/api/users", methods=["GET"])
@jwt_required()
def list_users():
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    users = User.query.order_by(User.id).all()
    return jsonify([u.to_dict() for u in users])


@app.route("/api/users", methods=["POST"])
@jwt_required()
def create_user():
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    data = request.get_json(silent=True) or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    if not username or not password:
        return jsonify({"error": "ユーザー名とパスワードは必須です"}), 400
    if User.query.filter_by(username=username).first():
        return jsonify({"error": "このユーザー名は既に使用されています"}), 409
    new_user = User(
        username=username,
        display_name=data.get("display_name", username),
        role=data.get("role", "user"),
        location=data.get("location"),
    )
    new_user.set_password(password)
    db.session.add(new_user)
    db.session.commit()
    return jsonify(new_user.to_dict()), 201


@app.route("/api/users/<int:uid>", methods=["PUT"])
@jwt_required()
def update_user(uid):
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    target = User.query.get_or_404(uid)
    data = request.get_json(silent=True) or {}

    if "display_name" in data:
        target.display_name = data["display_name"]
    if "role" in data:
        target.role = data["role"]
    if "location" in data:
        target.location = data["location"]
    if data.get("password"):
        target.set_password(data["password"])

    db.session.commit()
    return jsonify(target.to_dict())


@app.route("/api/users/<int:uid>", methods=["DELETE"])
@jwt_required()
def delete_user(uid):
    user = get_current_user()
    if user.role != "admin":
        return jsonify({"error": "管理者権限が必要です"}), 403
    if uid == user.id:
        return jsonify({"error": "自分自身は削除できません"}), 400
    target = User.query.get_or_404(uid)
    db.session.delete(target)
    db.session.commit()
    return jsonify({"message": "削除しました"}), 200


# ---------------------------------------------------------------------------
# Stats
# ---------------------------------------------------------------------------

@app.route("/api/stats", methods=["GET"])
@jwt_required()
def stats():
    user = get_current_user()
    q = Asset.query.filter_by(is_deleted=False)

    if user.role != "admin" and user.location != "サポートチーム":
        q = q.filter(Asset.location == user.location)

    total = q.count()
    by_category = q.with_entities(Asset.category, db.func.count(Asset.id)).group_by(Asset.category).all()
    by_location = q.with_entities(Asset.location, db.func.count(Asset.id)).group_by(Asset.location).all()
    maintenance_count = q.filter(Asset.maintenance_status == "有").count()

    return jsonify({
        "total": total,
        "by_category": {CATEGORY_LABELS.get(c, c): n for c, n in by_category},
        "by_location": dict(by_location),
        "maintenance_count": maintenance_count,
    })


# ---------------------------------------------------------------------------
# Health & Frontend
# ---------------------------------------------------------------------------

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "timestamp": datetime.utcnow().isoformat()})


@app.route("/")
def serve_frontend():
    return send_file("index.html")


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------

with app.app_context():
    db.create_all()
    seed_data()
    migrate_data()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
