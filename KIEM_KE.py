from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import os, datetime, qrcode

app = Flask(__name__)
app.secret_key = "kiemke_secret"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.xlsx")
CHECKIN_FILE = os.path.join(BASE_DIR, "checkin.csv")
STATIC_DIR = os.path.join(BASE_DIR, "static")
QR_FILE = os.path.join(STATIC_DIR, "qr_kiem_ke.png")

LINK = "https://kk-final-1.onrender.com"   # đổi đúng link Render của bạn

# ================= QR =================
os.makedirs(STATIC_DIR, exist_ok=True)
if not os.path.exists(QR_FILE):
    qrcode.make(LINK).save(QR_FILE)

# ================= LOAD DATA =================
def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        df.columns = df.columns.str.strip()
        return df
    return pd.DataFrame()

df = load_data()

# ================= LOGIN USER =================
@app.route("/", methods=["GET","POST"])
def login():
    if request.method == "POST":
        ma_nv = request.form["ma_nv"]
        mat_khau = request.form["mat_khau"]

        nv = df[
            (df["Ma_NV"].astype(str)==ma_nv) &
            (df["Mat_khau"].astype(str)==mat_khau)
        ]

        if nv.empty:
            return "❌ Sai mã NV hoặc mật khẩu"

        session["user"] = ma_nv
        return redirect("/user")

    return render_template("login.html")

# ================= USER HOME =================
@app.route("/user")
def user_home():
    if "user" not in session:
        return redirect("/")

    nv = df[df["Ma_NV"].astype(str)==session["user"]].iloc[0]

    trang_thai = "Chưa kiểm kê"
    if os.path.exists(CHECKIN_FILE):
        checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
        rows = checked[checked["Ma_NV"].astype(str)==session["user"]]
        if not rows.empty:
            trang_thai = rows.iloc[-1]["Trang_thai"]

    return render_template(
        "user_home.html",
        nv=nv,
        trang_thai=trang_thai
    )

# ================= USER CHECKIN =================
@app.route("/user/checkin", methods=["POST"])
def user_checkin():
    if "user" not in session:
        return redirect("/")

    ma_nv = session["user"]
    trang_thai = request.form.get("trang_thai", "Đang KK")

    nv = df[df["Ma_NV"].astype(str)==ma_nv].iloc[0]

    row = {
        "Ma_NV": nv["Ma_NV"],
        "Ho_ten": nv["Ho_ten"],
        "Bo_phan_KK": nv["Bo_phan_KK"],
        "Thoi_gian": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Trang_thai": trang_thai
    }

    if os.path.exists(CHECKIN_FILE):
        checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")

        # không cho ghi thêm nếu đã kết thúc
        done = checked[
            (checked["Ma_NV"].astype(str)==ma_nv) &
            (checked["Trang_thai"]=="Kết thúc KK")
        ]
        if not done.empty:
            return "⚠️ Nhân viên đã kết thúc kiểm kê"

        checked = pd.concat([checked, pd.DataFrame([row])])
    else:
        checked = pd.DataFrame([row])

    checked.to_csv(CHECKIN_FILE, index=False, encoding="utf-8-sig")
    return redirect("/user")

# ================= ADMIN LOGIN =================
@app.route("/admin", methods=["GET","POST"])
def admin_login():
    if request.method == "POST":
        if request.form["user"]=="admin" and request.form["pass"]=="admin123":
            session["admin"] = True
            return redirect("/admin/home")
        return "❌ Sai tài khoản admin"
    return render_template("admin_login.html")

# ================= ADMIN HOME =================
@app.route("/admin/home", methods=["GET","POST"])
def admin_home():
    if "admin" not in session:
        return redirect("/admin")

    global df
    if request.method == "POST":
        file = request.files["file"]
        file.save(DATA_FILE)
        df = load_data()

    return render_template("admin_home.html")

# ================= ADMIN DASHBOARD =================
@app.route("/admin/dashboard")
def admin_dashboard():
    if "admin" not in session:
        return redirect("/admin")

    if df.empty:
        return "❌ Chưa có dữ liệu nhân viên"

    # ===== Tổng NV theo Bộ phận KK =====
    total = (
        df.groupby("Bo_phan_KK")["Ma_NV"]
        .count()
        .reset_index(name="Tong")
    )

    # ===== Đọc file checkin =====
    if os.path.exists(CHECKIN_FILE):
        checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
    else:
        checked = pd.DataFrame(columns=["Ma_NV", "Bo_phan_KK", "Tinh_trang"])

    # ===== Đang KK =====
    dang_kk = (
        checked[checked["Tinh_trang"] == "Đang KK"]
        .groupby("Bo_phan_KK")["Ma_NV"]
        .count()
        .reset_index(name="Dang_KK")
    )

    # ===== Kết thúc KK =====
    ket_thuc = (
        checked[checked["Tinh_trang"] == "Kết thúc KK"]
        .groupby("Bo_phan_KK")["Ma_NV"]
        .count()
        .reset_index(name="Ket_thuc")
    )

    # ===== Gộp =====
    stat = (
        total
        .merge(dang_kk, on="Bo_phan_KK", how="left")
        .merge(ket_thuc, on="Bo_phan_KK", how="left")
        .fillna(0)
    )

    stat["Dang_KK"] = stat["Dang_KK"].astype(int)
    stat["Ket_thuc"] = stat["Ket_thuc"].astype(int)
    stat["Tien_do"] = (stat["Ket_thuc"] / stat["Tong"] * 100).round(1)

    return render_template(
        "admin_dashboard.html",
        stat=stat.to_dict(orient="records")
    )
# ================= EXPORT THEO BỘ PHẬN =================
@app.route("/admin/export/<bo_phan>")
def export_by_bo_phan(bo_phan):
    if "admin" not in session:
        return redirect("/admin")

    if not os.path.exists(CHECKIN_FILE):
        return "❌ Chưa có dữ liệu kiểm kê"

    dfc = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")

    df_bp = dfc[dfc["Bo_phan_KK"]==bo_phan]
    if df_bp.empty:
        return "Không có dữ liệu"

    path = os.path.join(BASE_DIR, f"kiem_ke_{bo_phan}.xlsx")
    df_bp.to_excel(path, index=False)

    return send_file(path, as_attachment=True)

# ================= RUN =================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",10000)))

