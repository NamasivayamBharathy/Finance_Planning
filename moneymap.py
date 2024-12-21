from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import financetool
import pandas as pd

app = Flask(__name__)

# Enable template auto-reloading
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Define the folder for generated files
GENERATED_FILES_FOLDER = 'static/generated_files'

def delete_existing_sheets(folder_path):
    """Deletes all Excel files in the specified folder."""
    try:
        for file in os.listdir(folder_path):
            if file.endswith('.xlsx'):
                file_path = os.path.join(folder_path, file)
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
    except Exception as e:
        print(f"Error while deleting files: {str(e)}")

@app.route('/')
def welcome():
    return render_template('welcome.html')

@app.route('/input', methods=['GET', 'POST'])
def input_page():
    if request.method == 'POST':
        try:
            username = request.form['user_name']
            inputs = {
                'initial_age': int(request.form['current_age']),
                'retirement_age': int(request.form['retirement_age']),
                'monthly_expenses': float(request.form['monthly_expenses']),
                'monthly_in_hand': float(request.form['monthly_salary']),
                'monthly_emi': float(request.form.get('monthly_emi', 0)),
                'initial_pf_corpus': float(request.form['pf_corpus']),
                'monthly_pf_contribution': float(request.form['pf_contribution']),
                'initial_current_corpus': float(request.form['investment_corpus']),
                'goals': []
            }

            # Collect goals dynamically
            for i in range(1, 6):
                goal_type = request.form.get(f'Goal{i}')
                if goal_type:
                    target_amount = float(request.form.get(f'Goal{i}_target', 0))
                    years_to_achieve = int(request.form.get(f'Goal{i}_years', 0))
                    inputs['goals'].append({
                        'type': goal_type,
                        'amount': target_amount,
                        'years': years_to_achieve
                    })

            delete_existing_sheets(GENERATED_FILES_FOLDER)

            result = financetool.main(data=inputs)

            if result['success']:
                output_filename = f"financial_report_{username}.xlsx"
                output_file_path = os.path.join(GENERATED_FILES_FOLDER, output_filename)

                if not os.path.exists(output_file_path):
                    os.rename(result['output_path'], output_file_path)

                # Read the "Summary" sheet
                df_summary = pd.read_excel(output_file_path, sheet_name="Summary")
                summary_html = df_summary.to_html(classes='table table-bordered', index=False)

                return render_template(
                    'result.html',
                    username=username,
                    output_path=output_filename,
                    summary_html=summary_html
                )
            else:
                return f"An error occurred: {result.get('error', result.get('message'))}", 500

        except Exception as e:
            return render_template('error.html', error_message=str(e)), 500
    else:
        return render_template('input_form.html')

@app.route('/retirement_plan/<filename>')
def retirement_plan(filename):
    try:
        file_path = os.path.join(GENERATED_FILES_FOLDER, filename)
        retirement_sheets = ["Retirement Plan", "Investment Allocation"]
        excel_data = {}

        for sheet_name in retirement_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            excel_data[sheet_name] = df.to_html(classes='table table-bordered', index=False)

        return render_template('retirement_plan.html', excel_data=excel_data)
    except Exception as e:
        return render_template('error.html', error_message=str(e)), 500

@app.route('/goals_plan/<filename>')
def goals_plan(filename):
    try:
        file_path = os.path.join(GENERATED_FILES_FOLDER, filename)
        all_sheets = pd.ExcelFile(file_path).sheet_names
        goals_sheets = [sheet for sheet in all_sheets if "Goal" in sheet]
        excel_data = {}

        for sheet_name in goals_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            excel_data[sheet_name] = df.to_html(classes='table table-bordered', index=False)

        return render_template('goals_plan.html', excel_data=excel_data)
    except Exception as e:
        return render_template('error.html', error_message=str(e)), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(GENERATED_FILES_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return f"Error: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
