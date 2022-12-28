from Utilities import *


class PageGenerator:
    def __init__(self, html_data, frame_data):
        if html_data is None or len(html_data) == 0:
            self._html_data = []
            self.frame_data = []
            self.action_options = {}
        else:
            self._html_data = html_data
            self.frame_data = frame_data

    def parse_html_data(self, base_url):
        if self._html_data is not None and len(self._html_data) != 0:
            arg_detail_info = ''
            for css_item in range(len(self._html_data)):
                row_content = self._html_data.iloc[css_item]
                name = re.sub(r'[^a-z0-9A-Z_]', '', row_content["name"], 0, flags=re.I)
                arg_name = name.strip() + '_locator'
                if name[:-1] == '_':
                    arg_name = name.strip() + 'locator'
                if row_content["frame"] != 'None':
                    arg_detail_info = arg_detail_info + '{0} = {1}'.format(arg_name.lower(), row_content["locator"]) \
                                      + '  # This items is under frame {0}'.format(row_content["frame"]) + '\n    '
                else:
                    arg_detail_info = arg_detail_info + '{0} = {1}'.format(arg_name.lower(), row_content["locator"]) \
                                      + '\n    '
            return self._output_html_data(base_url, arg_detail_info)

    def _output_html_data(self, base_url, arg_details):
        page_template_file = open(PAGE_TEMPLATE_PATH, 'r', encoding='utf-8')
        page_body_details_info = page_template_file.read()
        page_out_detail_info = page_body_details_info.format(base_url=base_url, arg_details=arg_details)

        return page_out_detail_info


