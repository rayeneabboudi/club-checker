import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import os

# 1. Initialize Firebase
if not firebase_admin._apps:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)

db = firestore.client()

def generate_reports():
    print("🔄 Fetching data from Firebase...")
    
    docs = db.collection("member_evaluations").stream()
    
    data_list = []
    for doc in docs:
        d = doc.to_dict()
        
        member = d.get('member_info', {})
        metrics = d.get('performance_metrics', {})
        feedback = d.get('feedback', {})
        meta = d.get('meta', {})
        deductions = d.get('deductions_log', d.get('deductions', {}))

        row = {
            "Department": member.get('department', 'General'),
            "Member Name": member.get('name', 'N/A'),
            "Evaluator": member.get('evaluator_name', 'N/A'),
            "Period": f"{member.get('period_start', '')} to {member.get('period_end', '')}",
            "Attendance": metrics.get('attendance', 0),
            "Tasks": metrics.get('task_execution', 0),
            "Initiative": metrics.get('initiative', 0),
            "Interaction": metrics.get('team_interaction', 0),
            "Deductions": deductions.get('total_penalty', deductions.get('total_points_lost', 0)),
            "Final Score": metrics.get('net_final_score', 0),
            "Rating": feedback.get('rating_label', 'N/A'),
            "Date Submitted": meta.get('submission_date', 'No Date')
        }
        data_list.append(row)

    if not data_list:
        print("❌ No data found in database.")
        return

    df = pd.DataFrame(data_list)
    output_folder = "Department_Reports"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    departments = df['Department'].unique()
    
    for dept in departments:
        clean_dept_name = str(dept).replace("/", "-").strip()
        filename = f"{output_folder}/{clean_dept_name}_Evaluations.xlsx"
        
        # --- NEW STYLING LOGIC START ---
        # Create a Pandas Excel writer using xlsxwriter as the engine
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        dept_df = df[df['Department'] == dept]
        
        # Write the dataframe to the sheet
        dept_df.to_excel(writer, sheet_name='Evaluations', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Evaluations']

        # Define specific styles
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#333333', # Dark gray (Bento style)
            'font_color': '#FFFFFF', # White text
            'border': 1
        })

        cell_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'align': 'left'
        })

        # Apply header formatting
        for col_num, value in enumerate(dept_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Set Column Widths for visibility
        worksheet.set_column('A:B', 18, cell_format) # Dept & Name
        worksheet.set_column('C:D', 22, cell_format) # Evaluator & Period
        worksheet.set_column('E:I', 12, cell_format) # Individual Scores
        worksheet.set_column('J:L', 15, cell_format) # Final, Rating, Date

        writer.close()
        # --- NEW STYLING LOGIC END ---
        
        print(f"✅ Stylized Report Created: {filename}")

if __name__ == "__main__":
    generate_reports()