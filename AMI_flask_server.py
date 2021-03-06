import logging
import random
import os
import report_generator
import status_changer
import download_playlist
import deadline_publisher
from dotenv import load_dotenv
from flask import Flask, request, session, redirect, render_template, send_file

load_dotenv('C:/Juice_Pipeline/config.env')
app = Flask(__name__)
app.secret_key = os.getenv('AMI_SECRET_KEY')  # dla obslugi sesji
host = os.getenv('AMI_FLASK_HOST')
port = os.getenv('AMI_FLASK_PORT')
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


@app.route('/download_playlist', methods=['POST'])
def download_playlists():
    post_dict = dict(request.form)
    playlist = download_playlist.PlaylistDownloader(post_dict)
    data = playlist.download()
    return render_template('download_playlist.html', data=data)


@app.route('/change_status', methods=['POST'])
def change_status():
    session['post_dict'] = dict(request.form)
    status = status_changer.StatusChanger(session['post_dict'])
    status.change_status()
    return 'status changed'


@app.route('/deadline_publisher', methods=['POST'])
def publish_deadline_to_sg():
    publisher = deadline_publisher.DeadlinePublisher(dict(request.form))
    publisher.publish()
    return 'render job submitted'


@app.route('/test')
def hello():
    app.logger.info('===========TEST DONE===========')
    return 'TEST'


if __name__ == "__main__":
    app.run(debug=True, host=host, port=port)
