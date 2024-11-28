from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

@app.route('/')
def show_winner():
    # Read the Excel file
    df = pd.read_excel('LuckyWinner.xlsx')
    
    # Get the last filled row
    last_row = df.iloc[-1].to_dict()
    
    # Remove the contact number from the dictionary
    if 'Contact No' in last_row:
        del last_row['Contact No']
    
    # Render the HTML template with the last row data
    return render_template('winner.html', winner=last_row)

if __name__ == '__main__':
    app.run(debug=True, port=8000)