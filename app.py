# Importing OS library
import os
from os.path import exists

# Importing psutil
import psutil

# Importing Bootstrap-Flask
from flask_bootstrap import Bootstrap5
# Importing flask module in the project is mandatory
# An object of Flask class is our WSGI application.
from flask import Flask, render_template, flash, redirect, url_for, send_file

# Importing WTForm, fields and validators
from flask_wtf import FlaskForm
from wtforms.fields import SubmitField, StringField
from flask_wtf.file import FileField, FileRequired, FileAllowed
from wtforms.validators import DataRequired

# Importing Pandas
import pandas as pd

# Flask constructor takes the name of
# current module (__name__) as argument.
app = Flask(__name__)

# Set Flask app secret key configuration
app.config['SECRET_KEY'] = 'my secrete key'

# Initializing Bootstrap-Flask
bootstrap = Bootstrap5(app)


class ExcelForm(FlaskForm):
    columns_order = StringField('Desired Columns Odering', validators=[DataRequired()],
                                render_kw={"class": "form-control", "placeholder": "0,1,2,3,4,5,.....[b for blank column]"})
    file = FileField('Excel File', validators=[
        FileRequired(),
        FileAllowed(['xlsx', 'xlsm'], 'Excel only!')
    ], render_kw={"class": "form-control"})
    submit = SubmitField('Submit', render_kw={"class": "form-control mb-3"})

# The route() function of the Flask class is a decorator,
# which tells the application which URL should call
# the associated function.


@app.route('/', methods=['GET', 'POST'])
# ‘/’ URL is bound with excel() function.
def excel():
    form = ExcelForm()
    try:
        if form.validate_on_submit():
            # Retrieve Excel file from form file field
            f = form.file.data

            # Retrieve Desired Columns Odering from the form
            # columns_order field
            ordering = form.columns_order.data

            # Split Desired Columns Odering into a list
            # where each character separated by comma is a
            # list item. splitting is done with
            # Python String split() Method
            odering_list = str(ordering).split(",")

            # Separate Desired Columns Ordering with data from
            # Desired blank columns ordering.
            # Recall that blanks are represented by the character
            # "b" case insensitive.
            odering_list2 = []  # Instantiate Variable (a list)
            # to hold Desired Columns Ordering
            # with data

            blank_order = []    # Instantiate Variable ( a list)
            # to hold Desired blank Columns
            # Ordering

            # Separate Desired column ordering as described above
            # by looping over odering_list pushing column order index
            # to respective category variable with Python list append()
            # Method
            for i in range(0, len(odering_list)):
                if odering_list[i] == 'b' or odering_list[i] == 'B':
                    blank_order.append(i)
                else:
                    odering_list2.append(int(odering_list[i]))

            # Read excel file to pandas DataFrame
            df = pd.read_excel(f.read(), engine='openpyxl')

            # Convert DataFrame to a dictionary with pandas.DataFrame.to_dict() Method
            df_dict = df.to_dict()

            # Extract DataFrame columns to a list with python dictionary .keys() Method
            cols = df_dict.keys()

            # Convert python dictionary to a 2-dimensional DataFrame based on extracted columns
            _2d_df = pd.DataFrame(df_dict, columns=cols)

            # Re-order the 2-dimensional DataFrame based on our Desired Columns Ordering
            re_ordered_2d_df = _2d_df.iloc[:, odering_list2]

            # Insert Blanks if desired
            if blank_order != []:
                for i in range(0, len(blank_order)):
                    re_ordered_2d_df.insert(
                        blank_order[i], "Blank_Column"+str(i), " ")

            # Create a panda object for writing to excel
            datatoexcel = pd.ExcelWriter('Data.xlsx')

            # write DataFrame to excel
            re_ordered_2d_df.to_excel(datatoexcel)

            # save the excel
            datatoexcel.save()

            # Notify User
            flash('Task completed successfully.')

        else:
            errors = form.errors
            print('Form data failed validations with errors: ', errors)
    except Exception as e:
        flash('An error occurred, check your file and input and try again.')
        print('An error occurred, check your file, and input and try again. ' + str(e))

    file_exists = exists('Data.xlsx')

    return render_template('index.html', form=form, file_exists=file_exists)


@app.route('/download', methods=['GET', 'POST'])
def download():
    path = 'Data.xlsx'

    # Check if file exist
    file_exists = exists('Data.xlsx')

    # Redirect if file does not exist
    if file_exists == False:
        flash('File does not exist')
        return redirect(url_for('excel'))

    return send_file(path, as_attachment=True)


@app.route('/delete', methods=['GET', 'POST'])
def delete():
    path = 'Data.xlsx'

    # Check if file exist
    file_exists = exists(path)

    # Redirect if file does not exist
    if file_exists == False:
        flash('File does not exist')
        return redirect(url_for('excel'))

    # Remove file
    try:
        os.remove(path)
    except:  # PermissionError: [WinError 32] The process cannot access the file because another process is using it
        for proc in psutil.process_iter():
            if proc.name() == 'python.exe':
                proc.kill()
        os.remove(path)

    return redirect(url_for('excel'))


# main driver function
if __name__ == '__main__':

    # run() method of Flask class runs the application
    # on the local development server.
    app.run(debug=True)
