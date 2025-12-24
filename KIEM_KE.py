from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import os, datetime

app = Flask(__name__)
app.secret_key = "kiemke_secret"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.xlsx")
CHECKIN_FILE = os.path.join(BASE_DIR, "checkin.csv")

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

        # admin
        if ma_nv == "admin" and mat_khau == "admin123":
            session.clear()
            session["admin"] = True
            return redirect("/dashboard")

        # user
        nv = df[
            (df["Ma_NV"].astype(str)==ma_nv) &
            (df["Mat_khau"].astype(str)==mat_khau)
        ]

        if nv.empty:
            return "❌ Sai tài khoản"

        session.clear()
        session["user"] = ma_nv
        return redirect("/dashboard")

    return render_template("login.html")

# ================= DASHBOARD CHUNG =================
@app.route("/dashboard")
def dashboard():
    # ===== ADMIN =====
    if "admin" in session:
        if not os.path.exists(CHECKIN_FILE):
            stat = []
        else:
            checked = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")

            total = df.groupby("Bo_phan_KK")["Ma_NV"].count().reset_index(name="Tong")

            dang = checked[checked["Trang_thai"]=="Đang KK"] \
                .groupby("Bo_phan_KK")["Ma_NV"].count().reset_index(name="Dang_KK")

            ket = checked[checked["Trang_thai"]=="Kết thúc KK"] \
                .groupby("Bo_phan_KK")["Ma_NV"].count().reset_index(name="Ket_thuc")

            stat = total.merge(dang, on="Bo_phan_KK", how="left") \
                        .merge(ket, on="Bo_phan_KK", how="left") \
                        .fillna(0)

            stat["Tien_do"] = (stat["Ket_thuc"]/stat["Tong"]*100).round(1)
            stat = stat.to_dict(orient="records")

        return render_template("dashboard.html", role="admin", stat=stat)

    # ===== USER =====
    if "user" in session:
        nv = df[df["Ma_NV"].astype(str)==session["user"]].iloc[0]

        trang_thai = "Chưa kiểm kê"
        if os.path.exists(CHECKIN_FILE):
            ck = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
            row = ck[ck["Ma_NV"].astype(str)==session["user"]]
            if not row.empty:
                trang_thai = row.iloc[-1]["Trang_thai"]

        return render_template(
            "dashboard.html",
            role="user",
            nv=nv,
            trang_thai=trang_thai
        )

    return redirect("/")

# ================= USER CHECKIN =================
@app.route("/checkin", methods=["POST"])
def checkin():
    if "user" not in session:
        return redirect("/")

    nv = df[df["Ma_NV"].astype(str)==session["user"]].iloc[0]
    trang_thai = request.form["trang_thai"]

    row = {
        "Ma_NV": nv["Ma_NV"],
        "Ho_ten": nv["Ho_ten"],
        "Bo_phan_KK": nv["Bo_phan_KK"],
        "Thoi_gian": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Trang_thai": trang_thai
    }

    if os.path.exists(CHECKIN_FILE):
        ck = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
        ck = pd.concat([ck, pd.DataFrame([row])])
    else:
        ck = pd.DataFrame([row])

    ck.to_csv(CHECKIN_FILE, index=False, encoding="utf-8-sig")
    return redirect("/dashboard")

# ================= EXPORT THEO BỘ PHẬN =================
@app.route("/export/<bo_phan>")
def export_bo_phan(bo_phan):
    if "admin" not in session:
        return redirect("/")

    ck = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
    out = ck[ck["Bo_phan_KK"]==bo_phan]

    if out.empty:
        return "❌ Không có dữ liệu"

    cols = ["Ma_NV","Ho_ten","Bo_phan_KK","Thoi_gian","Trang_thai"]
    out = out[cols]

    path = os.path.join(BASE_DIR, f"KQ_{bo_phan}.csv")
    out.to_csv(path, index=False, encoding="utf-8-sig")
    return send_file(path, as_attachment=True)

# ================= EXPORT ALL =================
@app.route("/export_all")
def export_all():
    if "admin" not in session:
        return redirect("/")

    ck = pd.read_csv(CHECKIN_FILE, encoding="utf-8-sig")
    cols = ["Ma_NV","Ho_ten","Bo_phan_KK","Thoi_gian","Trang_thai"]
    ck = ck[cols]

    path = os.path.join(BASE_DIR, "KQ_TAT_CA.csv")
    ck.to_csv(path, index=False, encoding="utf-8-sig")
    return send_file(path, as_attachment=True)

# ================= LOGOUT =================
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ================= RUN =================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",10000)))
