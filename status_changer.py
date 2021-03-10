import shotgun_api3
import os


sg_script_name = os.getenv('AMI_CHANGE_STATUS_NAME')
sg_token = os.getenv('AMI_CHANGE_STATUS_KEY')
sg_address = os.getenv('SG_ADDRESS')


class StatusChanger:
    def __init__(self, post_dict):
        self.__sg = shotgun_api3.Shotgun(sg_address, script_name=sg_script_name, api_key=sg_token)
        self.__post_dict = post_dict

    def change_status(self):
        page_id = self.__post_dict['page_id']
        page_id = int(page_id)
        current_page = self.__sg.find_one('Page', [['id', 'is', page_id]], ['name'])
        if self.is_page_accepted(current_page):
            self.change_entities_status()
            return True
        return False

    def change_entities_status(self):
        selected_entities_ids = self.__convert_to_list('ids')
        for entity_id in selected_entities_ids:
            status = self.__sg.find_one('Note', [['id', 'is', int(entity_id)]], ['sg_status_list'])
            status = status['sg_status_list']
            if status == 'clsd':
                status = 'ip'
            else:
                status = 'clsd'
            self.__sg.update('Note', int(entity_id), {'sg_status_list': status})
        return status

    def __convert_to_list(self, data):
        converted_list = self.__post_dict[data].split(',')
        return converted_list

    @staticmethod
    def is_page_accepted(page):
        if page['name'] == 'render' or 'Render':
            return True
        return False
