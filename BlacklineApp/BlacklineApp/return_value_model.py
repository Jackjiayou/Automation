class return_data(object):
    """docstring for return_data"""
    def __init__(self, is_success = False, msg='',continiu_run = True):
        self.is_success = is_success
        self.msg = msg
        self.continiu_run = continiu_run

class return_font_data(object):
    """docstring for return_data"""
    def __init__(self, is_success = True, msg='',title = 'Info',width=400):
        self.is_success = is_success
        self.msg = msg
        self.title = title
        self.width = width
