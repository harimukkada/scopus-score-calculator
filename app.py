from flask import Flask, render_template, request

app = Flask(__name__)

@app.route("/")
def home():
    return "Flask App is running!"

@app.route("/submit", methods=["POST"])
def submit():
    if request.method == "POST":
        scopus_author_id = request.form["scopus_author_id"]
        return f"Submitted Scopus Author ID: {scopus_author_id}"

if __name__ == "__main__":
    app.run(debug=True)
