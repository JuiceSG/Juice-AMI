import urllib.request
import os
import logging
import shotgun_api3
from dotenv import load_dotenv

load_dotenv('C:/Juice_Pipeline/config.env')
sg_project_location = os.getenv('SG_PROJECT_LOCATION')
sg_token = os.getenv('SG_AMI_DOWNLOAD_PLAYLIST_TOKEN')
sg_script_name = os.getenv('SG_AMI_DOWNLOAD_PLAYLIST_NAME')
sg_address = os.getenv('SG_ADDRESS')
logging.basicConfig(level=logging.DEBUG)


class PlaylistDownloader:
    def __init__(self, post_dict):
        self.__data = {
            'project_name': post_dict["project_name"]
        }
        self.__sg = shotgun_api3.Shotgun(sg_address, script_name=sg_script_name, api_key=sg_token)
        self.__post_dict = post_dict

    def download(self):
        playlists = self.__get_playlists()
        for playlist in playlists:
            versions = playlist['versions']
            download_location = self.__get_download_location(playlist)
            self.__download_versions(versions, download_location)
        return self.__data

    def __get_playlists(self):
        project_id = self.__post_dict['project_id']
        project_id = int(project_id)
        selected_playlists_ids = self.__get_selected_ids()
        project = self.__sg.find_one('Project', [['id', 'is', project_id]])
        filter = [
            ['project', 'is', project],
            ['id', 'in', selected_playlists_ids],
        ]
        fields = ['versions', 'sg_uploaded_movie', 'code']
        playlists = self.__sg.find('Playlist', filter, fields)
        return playlists

    def __get_selected_ids(self):
        converted_list = self.__post_dict['selected_ids'].split(',')
        converted_list = map(lambda x: int(x), converted_list)
        converted_list = list(converted_list)
        return converted_list

    def __get_download_location(self, playlist):
        project_name = self.__post_dict["project_name"]
        playlist_name = playlist['code']
        playlist_name = self.__check_name_sanity(playlist_name)
        download_location = '%s%s/playlist' % (sg_project_location, project_name)
        self.__data['playlist_folder'] = download_location.replace('//192.168.1.204/Dane', 'X:')
        self.__check_folder(download_location)
        download_location = '%s/%s' % (download_location, playlist_name)
        self.__check_folder(download_location)
        return download_location

    def __download_versions(self, versions, download_location):
        for version in versions:
            version_data = self.__sg.find_one('Version', [['id', 'is', version['id']]], ['sg_uploaded_movie'])
            link = version_data['sg_uploaded_movie']['url']
            file = version_data['sg_uploaded_movie']['name']
            file = '%s/%s' % (download_location, file)
            if not os.path.exists(file):
                urllib.request.urlretrieve(link, file)

    @staticmethod
    def __check_folder(folder):
        is_exist = os.path.exists(folder)
        if is_exist is False:
            os.mkdir(folder)

    @staticmethod
    def __check_name_sanity(name):
        illegal_characters = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in name:
            if char in illegal_characters:
                name = name.replace(char, '_')
        return name