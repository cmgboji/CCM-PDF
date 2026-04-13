from flask import Flask, request, render_template, send_file
import pandas as pd
from matplotlib import font_manager, rcParams
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
    font_path = os.path.join('static', 'fonts', 'Montserrat-Regular.ttf')
    font_manager.fontManager.addfont(font_path)
    rcParams['font.family'] = 'Montserrat'

    counts = (series.iloc[:, column_index]
                    .dropna()
                    .clip(1, 5)
                    .value_counts(normalize=True)
                    .reindex([1, 2, 3, 4, 5], fill_value=0) * 100)


    plt.figure(figsize=(8, 5))
    ax = sns.barplot(x=counts.index, y=counts.values, color=color)
    ax.set_xlabel('Score')
    ax.set_title(title, pad=14)
    ax.set_ylim(0, max(100, counts.max() + 8))
    ax.yaxis.set_visible(False)
    ax.tick_params(bottom=False)

    for spine in ['top', 'right', 'left']:
        ax.spines[spine].set_visible(False)
    ax.spines['bottom'].set_visible(True)

    for bar, pct in zip(ax.patches, counts.values):
        label_y = min(pct + 1, ax.get_ylim()[1] - 4)
        if pct > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                label_y,
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
    font_path = os.path.join('static', 'fonts', 'Montserrat-Regular.ttf')
    font_manager.fontManager.addfont(font_path)
    rcParams['font.family'] = 'Montserrat'

    counts = data.iloc[:,column].value_counts(normalize=True) * 100

    plt.figure(figsize=(6, 4))
    ax = sns.barplot(x=counts.values, y=counts.index, color=color)
    plt.ylabel(None)
    plt.title(title)
    plt.xlim(0, 100)
    plt.xticks(rotation=45)
    plt.box(False)
    ax.xaxis.set_visible(False)
    ax.tick_params(left=False)

    for bar, pct in zip(ax.patches, counts.values):
        if pct > 0:
            ax.text(
                bar.get_width() + 1,
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
        return render_template("error.html",
                               title="Uh oh, something went wrong!", 
                               message="There was an error reading the Excel file. Please ensure the file is properly formatted and try again.",
                               hint="Hint: Make sure the Excel file contains the sheets 'Raw Self Assessment' and 'Graduate Responses', and that they are formatted correctly.")
    
    try:
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
        df = df[df['Timestamp'].dt.year == int(year)]
        grad = grad[grad['Follow-Up Period'] == '12 Months']
    except Exception as e:
        return render_template("error.html",
                               title="Uh oh, something went wrong!",
                               message="There was a time-related error. Please make sure that you entered a valid year and that the 'Timestamp' column in your Excel file is properly formatted.",
                               hint="Hint: Double-check that the 'Timestamp' column in your Excel file contains valid date entries and that the year you entered matches the year of the timestamps in the data.")

    
    try:
        social = df.iloc[:, 4:8]
        auto = df.iloc[:, 8:11]
        finance = df.iloc[:, 11:16]
        trauma = df.iloc[:, 16:21]
        engagement = df.iloc[:, 21:27]

        housing_list = ['Stable/Permanent housing', 'Temporary housing', 'Staying with friends or family', 'Transitional program']
        grad.loc[~grad['What is your current housing situation?'].isin(housing_list), 'What is your current housing situation?'] = 'Other'

    except Exception as e:
        return render_template("error.html",
                               title="Uh oh, something went wrong!",
                               message="There was an error processing the data. Please ensure that the columns in your Excel file are structured correctly and try again.",
                               hint="Hint: Make sure that the 'Raw Self Assessment' sheet contains the expected columns for social, autonomy, finance, trauma, and engagement data, and that the 'Graduate Responses' sheet contains the expected columns for housing situation, employment status, and how graduates report to be doing.")
    
    teal = '#76b0af'
    coral = '#f3623d'
    yellow = '#f0d747'

    try:
        if len(grad) > 0:
            plot_paths = {
                "connected": plot_yag(social, 0, 'Feeling connected to others', yellow, "connected_plot"),
                "participation": plot_yag(social, 2, 'Actively participates in group discussions', yellow, "participation_plot"),
                "finance": plot_yag(finance, 1, 'Has secured or actively working towards securing employment', coral, "finance_plot"),
                "schedule": plot_yag(auto, 0, 'Creates daily schedules to meet responsibilites', coral, "schedule_plot"),
                "guidance": plot_yag(trauma, 2, 'Seeks guidance when facing challenges', teal, "guidance_plot"),
                "coping": plot_yag(trauma, 3, 'Uses and incorporates coping strategies', teal, "coping_plot"),
                "progress": plot_yag(engagement, 0, 'Reports feeling significant progress', yellow, "progress_plot"),
                "support": plot_yag(engagement, 1, 'Reports feeling supported by program and staff', yellow, "support_plot"),
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
        return render_template("error.html",
                               title="Uh oh, something went wrong!",
                               message="There was an error generating the plots. Please ensure that the data is structured correctly and try again.",
                               hint="Hint: Make sure that the columns in your Excel file are structured correctly for the social, autonomy, finance, trauma, and engagement data, and that there is appropriate data for the graduate responses if you are including those plots.")

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
        return render_template("error.html",
                               title="Uh oh, something went wrong!",
                               message="There was an error calculating the statistics. Please ensure that the data is structured correctly and try again.",
                               hint="Hint: Make sure entered the correct year and that there is appropriate data for the graduate responses if you are including those statistics.")

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
        return render_template("error.html",
                               title="Uh oh, something went wrong!",
                               message="There was an error generating the PDF report. Please ensure that the data is structured correctly and that all necessary plots and statistics were generated successfully.",
                               hint="Hint: This error could be caused by an issue with the HTML template rendering or the PDF generation process. Contact technical support to see if there is an issue with Render, the WeasyPrint library, HTML template, or CSS styles.")


if __name__ == "__main__":
    app.run(debug=True)