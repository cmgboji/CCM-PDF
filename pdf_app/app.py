from flask import Flask, request, render_template, send_file

app = Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/run', methods=['POST'])
def run():
    year = request.form.get('year')
    print("Received year:", year)

    # For now, return a simple response
    return f"PDF for the year {year} generated successfully!"

if __name__ == "__main__":
    app.run(debug=True)