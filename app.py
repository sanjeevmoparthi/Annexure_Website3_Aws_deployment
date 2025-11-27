import os
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from io import BytesIO
import annexure_functions as annex

# -------------------------------------------------------
#                FLASK APP CONFIGURATION
# -------------------------------------------------------

app = Flask(__name__, static_folder='static', template_folder='templates')

app.secret_key = "change_this_in_prod"

app.config['UPLOAD_FOLDER'] = "uploads"
app.config['OUTPUT_FOLDER'] = "output"
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024   # 500 MB

ALLOWED_EXT = {'xls', 'xlsx'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


# -------------------------------------------------------
#                    PAGE ROUTES
# -------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/next")
def next():
    return render_template("index2.html")

@app.route("/nextAgain")
def nextAgain():
    return render_template("index3.html")

@app.route("/index4")
def index4():
    return render_template("index4.html")


# -------------------------------------------------------
#                     MERGE PAGE
# -------------------------------------------------------

@app.route("/merge", methods=["POST"])
def merge_files():
    uploaded_files = request.files.getlist("files")

    if not uploaded_files:
        flash("Please upload at least one file", "danger")
        return redirect(url_for('upload_page'))

    all_dataframes = []

    for file in uploaded_files:
        if file.filename == "":
            continue

        base_name = os.path.basename(file.filename)
        branch_name = os.path.splitext(base_name)[0]

        save_path = os.path.join(app.config['UPLOAD_FOLDER'], base_name)
        file.save(save_path)

        df = pd.read_excel(save_path)
        df.insert(0, "Branch", branch_name)
        all_dataframes.append(df)

    combined_df = pd.concat(all_dataframes, ignore_index=True)

    output_path = os.path.join(app.config['OUTPUT_FOLDER'], "Combined_Data.xlsx")
    combined_df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)


# -------------------------------------------------------
#                   ANNEXURE DOWNLOAD ROUTES
# -------------------------------------------------------

@app.route('/download/1', methods=['POST'])
def download_annexure1():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index2'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx files allowed", "danger")
        return redirect(url_for('index2'))

    try:
        df = pd.read_excel(file)
        out_io = annex.annexure1_generate_excel_bytes(df)

        return send_file(out_io, as_attachment=True,
                         download_name="Annexure1_Vendor_Wise_Margin.xlsx")
    except Exception as e:
        flash(f"Error: {e}", "danger")
        return redirect(url_for('index2'))


@app.route('/download/2', methods=['POST'])
def download_annexure2():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index2'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx files allowed", "danger")
        return redirect(url_for('index2'))

    try:
        df = pd.read_excel(file)
        out_io = annex.annexure2_generate_excel_bytes(df)

        return send_file(out_io, as_attachment=True,
                         download_name="Annexure2_Brand_Wise_Margin.xlsx")
    except Exception as e:
        flash(f"Error: {e}", "danger")
        return redirect(url_for('index2'))


# -------------------------------------------------------
#                     ANNEXURE – 3
# -------------------------------------------------------
@app.route('/download/3', methods=['POST'])
def download_annexure3():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine="openpyxl")
        out_io = annex.annexure3_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure3_Brand_Wise_Sales.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-3: {str(e)}", "danger")
        return redirect(url_for('index'))


# -------------------------------------------------------
#                     ANNEXURE – 4
# -------------------------------------------------------
@app.route('/download/4', methods=['POST'])
def download_annexure4():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine="openpyxl")
        out_io = annex.annexure4_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure4_Product_Wise_Sales_Summary.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-4: {str(e)}", "danger")
        return redirect(url_for('index'))


# -------------------------------------------------------
#                     ANNEXURE – 5
# -------------------------------------------------------
@app.route('/download/5', methods=['POST'])
def download_annexure5():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine="openpyxl")
        out_io = annex.annexure5_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure5_Product_Category_Contribution.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-5: {str(e)}", "danger")
        return redirect(url_for('index'))





# -------------------------------------------------------
#                     ANNEXURE – 6
# -------------------------------------------------------
@app.route('/download/6', methods=['POST'])
def download_annexure6():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine="openpyxl")
        out_io = annex.annexure6_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure6_Profit_Below_10.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-6: {str(e)}", "danger")
        return redirect(url_for('index'))
    

# -------------------------------------------------------
#                     ANNEXURE – 8
# -------------------------------------------------------

@app.route('/download/8', methods=['POST'])
def download_annexure8():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file, engine="openpyxl")

        # Make sure your function returns BYTES, not BytesIO!!
        excel_bytes = annex.annexure8_generate_excel_bytes(df)

        # Convert BYTES → BytesIO object
        output_stream = BytesIO(excel_bytes)
        output_stream.seek(0)

        return send_file(
            output_stream,
            as_attachment=True,
            download_name="Annexure8_SP_LessThan_LCost.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-8: {str(e)}", "danger")
        return redirect(url_for('index'))
    
@app.route('/download/9', methods=['POST'])
def download_annexure9():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    try:
        df = pd.read_excel(file)
        excel_stream = annex.annexure9_generate_excel_bytes(df)
        excel_stream.seek(0)

        return send_file(
            excel_stream,
            as_attachment=True,
            download_name="Annexure9_Neither_Profit_Nor_Loss.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-9: {e}", "danger")
        return redirect(url_for('index'))

@app.route('/download/10', methods=['POST'])
def download_annexure10():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    try:
        # Read the uploaded file into DataFrame
        df = pd.read_excel(file)

        # Generate Excel file bytes using annexure10 function
        excel_stream = annex.annexure10_generate_excel_bytes(df)
        excel_stream.seek(0)

        return send_file(
            excel_stream,
            as_attachment=True,
            download_name="Annexure10_HighVendorLessProfit.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-10: {e}", "danger")
        return redirect(url_for('index'))
    
@app.route('/download/11', methods=['POST'])
def download_annexure11():

    closing_file = request.files.get('closing')
    sales_file = request.files.get('sales')
    ibts_file = request.files.get('ibts')

    if not all([closing_file, sales_file, ibts_file]):
        flash("❌ Upload all 3 files: Closing, Sales, IBTS", "danger")
        return redirect(url_for("index4"))

    try:
        closing_df = pd.read_excel(closing_file, skiprows=6)
        sales_df   = pd.read_excel(sales_file)
        ibts_df    = pd.read_excel(ibts_file)

        out_io = annex.annexure11_generate_excel_bytes(closing_df, sales_df, ibts_df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure11_NonMovement.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"❌ Annexure-11 Error: {str(e)}", "danger")
        return redirect(url_for("index4"))




# -------------------------------------------------------
#                     RUN APP
# -------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000,use_reloader = False)




