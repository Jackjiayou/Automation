import logging

def check_have_avande(list_name):
    is_success = False
    try:
        for name in list_name:
            if 'Ava' in name:
              is_success =  True
        return is_success
    except Exception as ex:
        logging.exception('check_have_avande :'+str(ex))
        return False

def get_upload_file_name(discription):
    '''get upload file name'''
    file_name = ''
    try:
        special_name = '/141100/236001-236499/271000'
        if special_name in discription:
            return discription

        have_space_country = ['Costa Rica','Czech Republic','Russian Federation','Saudi Arabia','South Africa','Trinidad,Tobago','United Kingdom']
        list_name = discription.split()
        list_number = discription.split('-')
        company_name = ''
        country_name =''

        if check_have_avande(list_name):
            company_name = 'Avanade_'
        else :
            company_name = 'Accenture_'

        if list_name[0]+' '+list_name[1]+' '+list_name[2] == 'United Arab Emirates':
            country_name = 'United Arab Emirates'
        else:
             temp_name=list_name[0]+' '+list_name[1]
             for name in have_space_country:
                if temp_name == name:
                    country_name = temp_name  
                    break
                else:
                    country_name = list_name[0]

        file_frist_name = country_name + company_name
        end_number_name = ''
        is_236 = list_number[len(list_number) - 1].strip().startswith('236')
        if is_236:
            end_number_name = '236XXX'
        else :
            end_number_name = list_number[len(list_number) - 1].strip()
        file_name = file_frist_name + end_number_name
        return file_name
    except Exception as ex:
        print(123)

discription = 'Trinidad and Tobago GRP - Accrued Inpat Taxes CBF - 236001 to 236499'
tet  = get_upload_file_name(discription)
print(123)