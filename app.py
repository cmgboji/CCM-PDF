from flask import Flask, request, render_template, send_file
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from weasyprint import HTML
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html")

def plot(series, column_index, title, color, filename):
    os.makedirs("static/plots", exist_ok=True)
    counts = (series.iloc[:, column_index]
                    .dropna()
                    .clip(1, 5)
                    .value_counts(normalize=True)
                    .reindex([1, 2, 3, 4, 5], fill_value=0) * 100)

    plt.figure(figsize=(8, 5))
    ax = sns.barplot(x=counts.index, y=counts.values, color=color)
    ax.set_xlabel('Score')
    ax.set_title(title)
    ax.set_ylim(0, 100)
    ax.yaxis.set_visible(False)

    for spine in ['top', 'right', 'left']:
        ax.spines[spine].set_visible(False)
    ax.spines['bottom'].set_visible(True)

    for bar, pct in zip(ax.patches, counts.values):
        if pct > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                pct + 1,
                f'{pct:.0f}%',
                ha='center',
                va='bottom'
            )

    plt.box(False)

    save_path = f"static/plots/{filename}.jpg"
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.close()

    return save_path

def get_stats(df, col):
    return int(round(df.iloc[:,col][df.iloc[:,col] > 3].count()/df.iloc[:,col].dropna().count(), 2) * 100)

# def output_pdf(plot_paths, stats):
#     # Placeholder for PDF generation logic
#     pass

@app.route('/run', methods=['POST'])
def run():
    year = request.form.get('year')
    workbook = request.files.get('workbook')
    try:
        df = pd.read_excel(workbook, sheet_name='Raw Self Assessment', engine='openpyxl')
    except Exception as e:
        print("Error reading Excel file:", e)
        return "There was an error processing the uploaded file. Please ensure it's a valid Excel workbook (.xlsx) with the correct sheet name 'Raw Self Assessment'."

    
    try:
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
        df = df[df['Timestamp'].dt.year == int(year)]
    except Exception as e:
        print("Error converting Timestamp:", e)
        return "There was an error converting the Timestamp column."

    
    try:
        social = df.iloc[:, 4:8]
        auto = df.iloc[:, 8:11]
        finance = df.iloc[:, 11:16]
        trauma = df.iloc[:, 16:21]
        engagement = df.iloc[:, 21:27]
    except Exception as e:
        print("Error slicing DataFrame columns:", e)
        return "There was an error slicing the DataFrame. Please ensure the columns are structured as expected."
    
    teal = '#76b0af'
    coral = '#f3623d'
    yellow = '#f0d747'

    try:
        plot_paths = {
            "connected": plot(social, 0, 'Feeling connected to others', yellow, "connected_plot"),
            "participation": plot(social, 2, 'Actively participating in group activities and discussions', yellow, "participation_plot"),
            "finance": plot(finance, 1, 'Has secured or actively working towards securing employment', coral, "finance_plot"),
            "schedule": plot(auto, 0, 'Creates and follows daily schedule to meet responsibilites', coral, "schedule_plot"),
            "guidance": plot(trauma, 2, 'Seeking guidance when facing challenges', teal, "guidance_plot"),
            "coping": plot(trauma, 3, 'Using and incorporating coping strategies', teal, "coping_plot"),
            "progress": plot(engagement, 0, 'Feeling significant progress', yellow, "progress_plot"),
            "support": plot(engagement, 1, 'Feeling supported by program and staff', yellow, "support_plot")
        }
    
    except Exception as e:
        print("Error generating plots:", e)
        return "There was an error generating the plots. Please ensure the data is structured correctly and try again."

    try:
        stats = {'s1': get_stats(social, 1), 
                'f1': get_stats(finance, 0), 
                'f2': get_stats(finance, 3), 
                't1': get_stats(trauma, 0), 
                't2': get_stats(trauma, 4), 
                'e1': get_stats(engagement, 5)}
        
    except Exception as e:
        print("Error calculating statistics:", e)
        return "There was an error calculating the statistics. Please ensure the data is structured correctly and try again."

    try:
        html = render_template(
            "report_template.html",
            base_url=request.host_url,
            year=year,
            connected=plot_paths["connected"],
            participation=plot_paths["participation"],
            finance=plot_paths["finance"],
            schedule=plot_paths["schedule"],
            guidance=plot_paths["guidance"],
            coping=plot_paths["coping"],
            progress=plot_paths["progress"],
            support=plot_paths["support"],
            s1=stats["s1"],
            f1=stats["f1"],
            f2=stats["f2"],
            t1=stats["t1"],
            t2=stats["t2"],
            e1=stats["e1"],
        )

        
        pdf_path = f"static/{year} Impact Report.pdf"
        base_dir = os.path.abspath(os.path.dirname(__file__))
        HTML(string=html, base_url=base_dir).write_pdf(
            pdf_path,
            stylesheets=None,
            presentational_hints=True
        )
        return send_file(pdf_path, as_attachment=True)
    
    except Exception as e:
        print("Error generating PDF:", e)
        return "There was an error generating the PDF. Please ensure the template is structured correctly and try again."


if __name__ == "__main__":
    app.run(debug=True)