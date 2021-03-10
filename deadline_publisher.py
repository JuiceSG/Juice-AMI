import urllib.request
import os
import logging
import shotgun_api3
from dotenv import load_dotenv

load_dotenv('C:/Juice_Pipeline/config.env')
sg_token = os.getenv('AMI_DEADLINE_PUBLISHER_KEY')
sg_script_name = os.getenv('AMI_DEADLINE_PUBLISHER_NAME')
sg_address = os.getenv('SG_ADDRESS')
logging.basicConfig(level=logging.DEBUG)


class DeadlinePublisher:
    def __init__(self, post_dict):
        self.__data = post_dict
        self.__sg = shotgun_api3.Shotgun(sg_address, script_name=sg_script_name, api_key=sg_token)

    def publish(self):
        project_id = int(self.__data['project_id'])
        task_id = int(self.__data['task_id'])
        job_name = self.__data['job_name']
        version = job_name.split('_')[-1]
        version = int(version.replace('v', ''))
        project = self.__sg.find_one('Project', [['id', 'is', project_id]])
        task = self.__sg.find_one('Task', [['id', 'is', task_id]], ['entity'])
        entity = task['entity']
        path = self.get_path(self.__data['render_output'])
        published_file_type = self.__sg.find_one('PublishedFileType', [['code', 'is', 'Rendered Image']])
        path = {'local_path': path}
        data = {
            'code': job_name,
            'entity': entity,
            'version_number': version,
            'task': task,
            'path': path,
            'project': project,
            'sg_status_list': 'cmpt',
            'name': 'test123',
            'published_file_type': published_file_type
        }
        result = self.__sg.create('PublishedFile', data)
        return result

    def get_path(self, path):
        layer_name = self.__data
        files_list = os.listdir(path)
        file_name = [i for i in files_list if i.split('.')[-3].lower() == 'beauty']
        if layer_name == 'None':
            file_name = [i for i in files_list if i.split('.')[-3].lower() == 'beauty']
        else:
            file_name = [i for i in files_list if i.split('.')[-3].lower() == 'beauty' and i.split('.')[-4] == layer_name]
        if file_name:
            file_name = file_name[0]
        else:
            file_name = files_list[0]
        seq_name = self.convert_to_seq_name(file_name)
        path = os.path.join(path, seq_name)
        return path

    def convert_to_seq_name(self, file_name):
        seq_name = file_name.split('.')
        seq_format = self.get_seq_format(seq_name[-2])
        seq_name[-2] = seq_format
        seq_name = '.'.join(seq_name)
        return seq_name

    @staticmethod
    def get_seq_format(seq):
        seq_format = len(seq)
        seq_format = '%{:02}d'.format(seq_format)
        return seq_format