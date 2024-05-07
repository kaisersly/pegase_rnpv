from tempfile import NamedTemporaryFile
from flask import Flask, request, render_template, send_file

from parse_pv import parse_pv
from merge import merge_pv

app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/")
def action():
    action = request.form["action"]
    if not request.files["pv"]:
        return "Veuillez renseigner un PV .csv"
    pv_csv = NamedTemporaryFile()
    request.files["pv"].save(pv_csv.name)
    pv = parse_pv(pv_csv.name)
    if action == "list-merge-fields":
        merge_fields = []
        for field in pv.fixed_merge_fields.keys():
            merge_fields.append((field, "GLOBAL"))
        for field in pv.etudiants[0].keys():
            merge_fields.append((field, "LIGNE ETUDIANT"))
        return render_template("list_merge_fields.html", merge_fields=merge_fields)
    elif action == "merge":
        if not request.files["template"]:
            return "Veuillez renseigner une template .xlsx"
        template = NamedTemporaryFile(suffix=".xlsx")
        request.files["template"].save(template.name)
        merge = NamedTemporaryFile(suffix=".xlsx", delete=False)
        merge_pv(merge.name, pv, template.name)
        return send_file(
            merge.name, download_name=f"{pv.fixed_merge_fields['FORMATION_CODE']}.xlsx"
        )


if __name__ == "__main__":
    app.run(debug=True)
