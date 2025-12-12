# password_server.py
from getpass import getpass
from flask import Flask, jsonify

app = Flask(__name__)

SAP_USER = "Amellag"   # pon tu usuario aquí
SAP_PASS = None               # se llenará solo en memoria

@app.route("/sap_credentials")
def sap_credentials():
    global SAP_PASS
    if SAP_PASS is None:
        # Te pedirá la clave solo la primera vez
        SAP_PASS = getpass("Ingresa tu contraseña SAP (no se verá en pantalla): ")
    return jsonify({
        "user": SAP_USER,
        "password": SAP_PASS
    })

if __name__ == "__main__":
    # Solo escucha en tu PC (localhost)
    app.run(host="127.0.0.1", port=5001)
