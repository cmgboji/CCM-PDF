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

def plot_yag(series, column_index, title, color, filename):
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

def stats_yag(df, col):
    return int(round(df.iloc[:,col][df.iloc[:,col] > 3].count()/df.iloc[:,col].dropna().count(), 2) * 100)

def plot_imp(data, column, title, color, filename):
    os.makedirs("static/plots", exist_ok=True)
    counts = data.iloc[:,column].value_counts(normalize=True) * 100

    plt.figure(figsize=(6, 4))
    ax = sns.barplot(x=counts.values, y=counts.index, color=color)
    plt.ylabel(None)
    plt.title(title)
    plt.xlim(0, 100)
    plt.xticks(rotation=45)
    plt.box(False)
    ax.xaxis.set_visible(False)

    for bar, pct in zip(ax.patches, counts.values):
        if pct > 0:
            ax.text(
                bar.get_width() + .2,
                bar.get_y() + bar.get_height() / 2,
                f'{pct:.0f}%',
                va='center'
            )

    save_path = f"static/plots/{filename}.jpg"
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.close()
    return save_path

def stats_imp(df, col):
    return int(round(((df.iloc[:,col][df.iloc[:,col] == "Yes"].count() / len(df)) * 100), 2))

@app.route('/run', methods=['POST'])
def run():
    year = request.form.get('year')
    workbook = request.files.get('workbook')
    try:
        df = pd.read_excel(workbook, sheet_name='Raw Self Assessment', engine='openpyxl')
        grad = pd.read_excel(workbook, sheet_name='Graduate Responses', engine='openpyxl')
    except Exception as e:
        print("Error reading Excel file:", e)
        return "There was an error processing the uploaded file. Please ensure it's a valid Excel workbook (.xlsx) with the correct sheet name 'Raw Self Assessment'."

    
    try:
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
        df = df[df['Timestamp'].dt.year == int(year)]
        grad = grad[grad['Follow-Up Period'] == '12 Months']
    except Exception as e:
        print("Error converting Timestamp:", e)
        return "There was a time-related error. Please make sure that you entered a valid year and that the 'Timestamp' column in your Excel file is properly formatted as dates."

    
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
        if len(grad) > 0:
            plot_paths = {
                "connected": plot_yag(social, 0, 'Feeling connected to others', yellow, "connected_plot"),
                "participation": plot_yag(social, 2, 'Actively participating in group activities and discussions', yellow, "participation_plot"),
                "finance": plot_yag(finance, 1, 'Has secured or actively working towards securing employment', coral, "finance_plot"),
                "schedule": plot_yag(auto, 0, 'Creates and follows daily schedule to meet responsibilites', coral, "schedule_plot"),
                "guidance": plot_yag(trauma, 2, 'Seeking guidance when facing challenges', teal, "guidance_plot"),
                "coping": plot_yag(trauma, 3, 'Using and incorporating coping strategies', teal, "coping_plot"),
                "progress": plot_yag(engagement, 0, 'Feeling significant progress', yellow, "progress_plot"),
                "support": plot_yag(engagement, 1, 'Feeling supported by program and staff', yellow, "support_plot"),
                "housing": plot_imp(grad, 5, 'Current Housing Situation of Graduates', coral, "housing_plot"),
                "employment": plot_imp(grad, 6, 'Employment Status of Graduates', teal, "employment_plot"),
                "doing": plot_imp(grad, 11, 'How Graduates Report to be Doing', yellow, "doing_plot")
            }
        else:
            plot_paths = {
                "connected": plot_yag(social, 0, 'Feeling connected to others', yellow, "connected_plot"),
                "participation": plot_yag(social, 2, 'Actively participating in group activities and discussions', yellow, "participation_plot"),
                "finance": plot_yag(finance, 1, 'Has secured or actively working towards securing employment', coral, "finance_plot"),
                "schedule": plot_yag(auto, 0, 'Creates and follows daily schedule to meet responsibilites', coral, "schedule_plot"),
                "guidance": plot_yag(trauma, 2, 'Seeking guidance when facing challenges', teal, "guidance_plot"),
                "coping": plot_yag(trauma, 3, 'Using and incorporating coping strategies', teal, "coping_plot"),
                "progress": plot_yag(engagement, 0, 'Feeling significant progress', yellow, "progress_plot"),
                "support": plot_yag(engagement, 1, 'Feeling supported by program and staff', yellow, "support_plot"),
             }
    except Exception as e:
        print("Error generating plots:", e)
        return "There was an error generating the plots. Please ensure the data is structured correctly and try again."

    try:
        if len(grad) > 0:
            stats = {'s1': stats_yag(social, 1), 
                    'f1': stats_yag(finance, 0), 
                    'f2': stats_yag(finance, 3), 
                    't1': stats_yag(trauma, 0), 
                    't2': stats_yag(trauma, 4), 
                    'e1': stats_yag(engagement, 5),
                    'sob': stats_imp(grad, 8),
                    'sup': stats_imp(grad, 10),
                    'con': stats_imp(grad, 13)
                    }
        else:
            stats = {'s1': stats_yag(social, 1), 
                    'f1': stats_yag(finance, 0), 
                    'f2': stats_yag(finance, 3), 
                    't1': stats_yag(trauma, 0), 
                    't2': stats_yag(trauma, 4), 
                    'e1': stats_yag(engagement, 5)
                    }
        
    except Exception as e:
        print("Error calculating statistics:", e)
        return "There was an error calculating the statistics. Please ensure the data is structured correctly and try again."

    try:
        grad_data = len(grad) > 0
        if grad_data:
            html = render_template(
                "report_template.html",
                base_url=request.host_url,
                year=year,
                grad_data=grad_data,
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
                housing=plot_paths["housing"],
                employment=plot_paths["employment"],
                doing=plot_paths["doing"],
                sob=stats["sob"],
                sup=stats["sup"],
                con=stats["con"]

            )
        else:
            html = render_template(
                "report_template.html",
                base_url=request.host_url,
                year=year,
                grad_data=grad_data,
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
                e1=stats["e1"]
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