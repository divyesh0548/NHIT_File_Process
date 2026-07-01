# from flask import Flask, render_template, request, jsonify, send_file
# import pandas as pd
# import os
# import tempfile
# import time

# app = Flask(__name__)

# UPLOAD_FOLDER = tempfile.gettempdir()
# header_map = {}
# uploaded_files = []

# # =========================
# # CSV SAFE READER
# # =========================
# def read_csv_safe(file_path):
#     with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
#         lines = f.readlines()

#     rows = [line.strip().split(",") for line in lines]
#     header = [col.replace("'", "").strip() for col in rows[0]]
#     data = rows[1:]

#     fixed_data = []
#     for row in data:
#         if len(row) > len(header):
#             row = row[:len(header)]
#         elif len(row) < len(header):
#             row = row + [""] * (len(header) - len(row))
#         fixed_data.append(row)

#     df = pd.DataFrame(fixed_data, columns=header)
#     return df

# # =========================
# # CSV READER READ TIME (COLUMN)
# # =========================
# def apply_reader_read_time_csv(df):
#     if "Reader Read time" in df.columns:
#         start_idx = df.columns.get_loc("Reader Read time")
#         df = df.iloc[:, start_idx:]
#     return df

# # =========================
# # EXCEL (2 SHEETS, HEADER ROW BY "Reader Read time")
# # =========================
# def read_excel_two_sheets(file_path):
#     xls = pd.ExcelFile(file_path)
#     sheets = xls.sheet_names

#     dfs = []

#     for sheet in sheets[:2]:  # only first 2 sheets
#         raw_df = pd.read_excel(file_path, sheet_name=sheet, header=None)

#         header_row_idx = None

#         for i, row in raw_df.iterrows():
#             if row.astype(str).str.contains("Reader Read time", case=False).any():
#                 header_row_idx = i
#                 break

#         if header_row_idx is None:
#             continue

#         header = raw_df.iloc[header_row_idx]
#         data = raw_df.iloc[header_row_idx + 1:]

#         data.columns = header
#         data = data.reset_index(drop=True)

#         dfs.append(data)

#     if dfs:
#         return pd.concat(dfs, ignore_index=True)
#     else:
#         return pd.DataFrame()

# # =========================
# # HEADER MAPPING
# # =========================
# def apply_header_mapping(df):
#     # rename mapped columns
#     df = df.rename(columns=lambda c: header_map.get(c, c))
#     return df

# # =========================
# # ROUTES
# # =========================
# @app.route("/")
# def index():
#     return render_template("index.html")

# @app.route("/upload-mapping", methods=["POST"])
# def upload_mapping():
#     global header_map, uploaded_files
#     uploaded_files = []  # reset files when new mapping uploaded

#     file = request.files["file"]
#     path = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(path)

#     mapping_df = pd.read_excel(path)
#     header_map = dict(zip(mapping_df["Replace From"], mapping_df["Replace To"]))

#     return "OK"

# @app.route("/upload-file", methods=["POST"])
# def upload_file():
#     file = request.files["file"]
#     path = os.path.join(UPLOAD_FOLDER, file.filename)
#     file.save(path)
#     uploaded_files.append(path)

#     return jsonify({"filename": file.filename})

# @app.route("/process")
# def process_files():
#     start_time = time.time()

#     combined_data = []

#     for path in uploaded_files:
#         if path.lower().endswith(".csv"):
#             df = read_csv_safe(path)
#             df = apply_reader_read_time_csv(df)
#         else:
#             df = read_excel_two_sheets(path)

#         df = apply_header_mapping(df)
#         combined_data.append(df)

#     final_df = pd.concat(combined_data, ignore_index=True)

#     # FILTER: Lane No <> blank or null
#     if "Lane No" in final_df.columns:
#         final_df = final_df[
#             final_df["Lane No"].notna() & (final_df["Lane No"].astype(str).str.strip() != "")
#         ]

#     output_path = os.path.join(UPLOAD_FOLDER, "combined_output.csv")
#     final_df.to_csv(output_path, index=False)

#     # Calculate elapsed time
#     elapsed_seconds = int(time.time() - start_time)
#     minutes, seconds = divmod(elapsed_seconds, 60)
#     elapsed_str = f"Processing completed in {minutes} min {seconds} sec"

#     return jsonify({
#         "message": elapsed_str,
#         "download_url": "/download-output"
#     })

# @app.route("/download-output")
# def download_output():
#     output_path = os.path.join(UPLOAD_FOLDER, "combined_output.csv")
#     return send_file(output_path, as_attachment=True)

# if __name__ == "__main__":
#     print("🚀 Server started at http://127.0.0.1:5000")
#     app.run(debug=True)








from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import tempfile
import time

app = Flask(__name__)

UPLOAD_FOLDER = tempfile.gettempdir()
header_map = {}
uploaded_files = []

# =========================
# CSV SAFE READER
# =========================
def read_csv_safe(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    rows = [line.strip().split(",") for line in lines]
    header = [col.replace("'", "").strip() for col in rows[0]]
    data = rows[1:]

    fixed_data = []
    for row in data:
        if len(row) > len(header):
            row = row[:len(header)]
        elif len(row) < len(header):
            row = row + [""] * (len(header) - len(row))
        fixed_data.append(row)

    df = pd.DataFrame(fixed_data, columns=header)
    return df

# =========================
# CSV READER READ TIME (COLUMN)
# =========================
def apply_reader_read_time_csv(df):
    if "Reader Read time" in df.columns:
        start_idx = df.columns.get_loc("Reader Read time")
        df = df.iloc[:, start_idx:]
    return df

# =========================
# EXCEL PROCESSING (FIRST 20 ROWS TO FIND HEADER)
# =========================
def read_excel_two_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names[:2]  # only first 2 sheets
    dfs = []

    for sheet in sheets:
        # Step 1: Read first 20 rows to find header
        preview_df = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=20)
        header_row_idx = None
        for i, row in preview_df.iterrows():
            if row.astype(str).str.contains("Reader Read time", case=False).any():
                header_row_idx = i
                break

        if header_row_idx is None:
            continue  # skip sheet if header not found

        # Step 2: Read full sheet from header row
        data = pd.read_excel(file_path, sheet_name=sheet, header=header_row_idx)

        # Step 3: Apply header mapping immediately
        data = apply_header_mapping(data)

        # Step 4: Filter early (Lane No)
        if "Lane No" in data.columns:
            data = data[data["Lane No"].notna() & (data["Lane No"].astype(str).str.strip() != "")]

        dfs.append(data)

    if dfs:
        return pd.concat(dfs, ignore_index=True)
    else:
        return pd.DataFrame()

# =========================
# HEADER MAPPING
# =========================
def apply_header_mapping(df):
    df = df.rename(columns=lambda c: header_map.get(c, c))
    return df

# =========================
# ROUTES
# =========================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload-mapping", methods=["POST"])
def upload_mapping():
    global header_map, uploaded_files
    uploaded_files = []  # reset files when new mapping uploaded

    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    mapping_df = pd.read_excel(path)
    header_map = dict(zip(mapping_df["Replace From"], mapping_df["Replace To"]))

    return "OK"

@app.route("/upload-file", methods=["POST"])
def upload_file():
    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)
    uploaded_files.append(path)

    return jsonify({"filename": file.filename})

@app.route("/process")
def process_files():
    start_time = time.time()
    combined_data = []

    for path in uploaded_files:
        if path.lower().endswith(".csv"):
            df = read_csv_safe(path)
            df = apply_reader_read_time_csv(df)
            df = apply_header_mapping(df)
            # Early filter
            if "Lane No" in df.columns:
                df = df[df["Lane No"].notna() & (df["Lane No"].astype(str).str.strip() != "")]
        else:
            df = read_excel_two_sheets(path)

        combined_data.append(df)

    final_df = pd.concat(combined_data, ignore_index=True) if combined_data else pd.DataFrame()

    output_path = os.path.join(UPLOAD_FOLDER, "combined_output.csv")
    final_df.to_csv(output_path, index=False)

    elapsed_seconds = int(time.time() - start_time)
    minutes, seconds = divmod(elapsed_seconds, 60)
    elapsed_str = f"Processing completed in {minutes} min {seconds} sec"

    return jsonify({
        "message": elapsed_str,
        "download_url": "/download-output"
    })

@app.route("/download-output")
def download_output():
    output_path = os.path.join(UPLOAD_FOLDER, "combined_output.csv")
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    print("🚀 Server started at http://127.0.0.1:5000")
    app.run(debug=False, host="0.0.0.0", port=5000)