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

LINK = "https://nhung-kk-v01.onrender.com/"   # đổi đúng link Render

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

# ================= LOGIN =================
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
            return "❌ Sai tài khoản hoặc mật khẩu"

        nv = nv.iloc[0]
        session["user"] = str(nv["Ma_NV"])
        session["role"] = nv.get("Role","user")

        return redirect("/dashboard")

    return render_template("login.html")

# ================= DASHBOARD ROUTER =================
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect("/")
    if session["role"] == "admin":
        return redirect("/dashboard/admin")
    return redirect("/dashboard/user")

# ================= USER DASHBOARD =================
@app.route("/dashboard/user")
def dashboard_user():
    if "user" not in session or session["role"]!="user":
        return redirect("/")

    nv = df[df["Ma_NV"].astype(str)==session["user"]].iloc[0]

    trang_thai = "Chưa kiểm kê"
    if os.path.exists(CHECKIN_FILE):
        checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
        rows = checked[checked["Ma_NV"].astype(str)==session["user"]]
        if not rows.empty:
            trang_thai = rows.iloc[-1]["Trang_thai"]

    return render_template(
        "dashboard.html",
        role="user",
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
        checked = pd.concat([checked, pd.DataFrame([row])])
    else:
        checked = pd.DataFrame([row])

    checked.to_csv(CHECKIN_FILE, index=False, encoding="utf-8-sig")
    return redirect("/dashboard/user")

# ================= ADMIN DASHBOARD =================
@app.route("/dashboard/admin")
def dashboard_admin():
    if "user" not in session or session["role"]!="admin":
        return redirect("/")

    # Tổng nhân viên theo bộ phận
    total = df.groupby("Bo_phan_KK")["Ma_NV"] \
              .count().reset_index(name="Tong")

    # Nếu chưa có file checkin
    if not os.path.exists(CHECKIN_FILE):
        stat = total.copy()
        stat["Dang_KK"] = 0
        stat["Dang_cheo"] = 0
        stat["Ket_thuc"] = 0
        stat["Tien_do"] = 0
        return render_template(
            "dashboard.html",
            role="admin",
            stat=stat.to_dict(orient="records")
        )

    checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")

    # Tránh lỗi file cũ
    if "Trang_thai" not in checked.columns:
        checked["Trang_thai"] = "Kết thúc KK"

    # ===== THỐNG KÊ =====
    dang = checked[checked["Trang_thai"]=="Đang KK"] \
        .groupby("Bo_phan_KK")["Ma_NV"] \
        .count().reset_index(name="Dang_KK")

    cheo = checked[checked["Trang_thai"]=="Đang check chéo"] \
        .groupby("Bo_phan_KK")["Ma_NV"] \
        .count().reset_index(name="Dang_cheo")

    ket = checked[checked["Trang_thai"]=="Kết thúc KK"] \
        .groupby("Bo_phan_KK")["Ma_NV"] \
        .count().reset_index(name="Ket_thuc")

    # Gộp
    stat = total \
        .merge(dang, on="Bo_phan_KK", how="left") \
        .merge(cheo, on="Bo_phan_KK", how="left") \
        .merge(ket, on="Bo_phan_KK", how="left") \
        .fillna(0)

    stat["Dang_KK"] = stat["Dang_KK"].astype(int)
    stat["Dang_cheo"] = stat["Dang_cheo"].astype(int)
    stat["Ket_thuc"] = stat["Ket_thuc"].astype(int)

    stat["Tien_do"] = (stat["Ket_thuc"] / stat["Tong"] * 100).round(1)

    return render_template(
        "dashboard.html",
        role="admin",
        stat=stat.to_dict(orient="records")
    )

# ================= EXPORT THEO BỘ PHẬN =================
@app.route("/admin/export/<bo_phan>")
def export_bo_phan(bo_phan):
    if session.get("role")!="admin":
        return redirect("/")

    df_export = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
    df_export = df_export[df_export["Bo_phan_KK"]==bo_phan]

    file_path = f"KQ_{bo_phan}.csv"
    df_export.to_csv(file_path, index=False, encoding="utf-8-sig")
    return send_file(file_path, as_attachment=True)

# ================= EXPORT ALL =================
@app.route("/admin/export_all")
def export_all():
    if session.get("role")!="admin":
        return redirect("/")

    return send_file(CHECKIN_FILE, as_attachment=True)
# ================= LOGOUT =================
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ================= RUN =================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",10000)))




