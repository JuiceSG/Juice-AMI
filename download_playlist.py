import urllib.request
import os
import logging
import shotgun_api3
from dotenv import load_dotenv

load_dotenv('C:/Juice_Pipeline/config.env')
sg_token = os.getenv('SG_AMI_FLASK_TOKEN')
sg_address = os.getenv('SG_ADDRESS')
sg_script_name = os.getenv('SG_AMI_FLASK_NAME')
logging.basicConfig(level=logging.DEBUG)


class PlaylistDownloader:
    def __init__(self, post_dict):
        project_id = post_dict['project_id']
        selected_playlists_ids = self.__convert_to_list('ids')
        self.__sg = shotgun_api3.Shotgun(sg_address, script_name=sg_script_name, api_key=sg_token)
        self.__post_dict = post_dict
        self.__project = self.__sg.find_one('Project', [['id', 'is', project_id]])
        self.__playlists = self.__sg.find('Playlist',
                                          [
                                              ['project', 'is', self.__project],
                                              ['id', 'in', selected_playlists_ids]
                                          ],
                                          ['versions', 'sg_uploaded_movie', 'code']
                                          )

    def __convert_to_list(self, data):
        converted_list = self.__post_dict[data].split(',')
        return converted_list

    def download(self):
        for playlist in self.__playlists:
            versions = playlist['versions']
            playlist_name = playlist['code']
            playlist_name = playlist_name.replace('/', '_')
