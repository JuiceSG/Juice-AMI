import logging
import random
import os
import report_generator
from dotenv import load_dotenv
from flask import Flask, request, session, redirect, render_template, send_file

load_dotenv('C:/Juice_Pipeline/config.env')
app = Flask(__name__)
app.secret_key = os.getenv('AG_AMI_SECRET_KEY')   # dla obslugi sesji
host = os.getenv('SG_AMI_FLASK_HOST')
port = os.getenv('SG_AMI_FLASK_PORT')
logging.basicConfig(level=logging.DEBUG)


@app.route('/generate_report', methods=['POST'])
def report_generator_redirection():
    rand_int = random.getrandbits(128)
    redirected_address = ('/generate_report/%s' % rand_int)
    session['post_dict'] = dict(request.form)  # zwraca wynik przesylania mtedoy POST w postaci slownika
    return redirect(redirected_address)


@app.route('/generate_report/<redirected_address>')
def generate_report(redirected_address):
    if 'post_dict' in session:
        post_dict = session['post_dict']
    else:
        post_dict = None
        return 'Session problem, POST: %s' % post_dict
    new_report = report_generator.ReportGenerator(post_dict)
    document_title = new_report.title
    new_report = new_report.generate()
    return send_file(new_report, as_attachment=True, attachment_filename=document_title)


if __name__ == "__main__":
    app.run(debug=True, host=host, port=port)
