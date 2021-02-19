import xlsxwriter
import itertools
import string

# get excel column letters
excel_columns=list(itertools.chain(string.ascii_uppercase, (''.join(pair) for pair in itertools.product(string.ascii_uppercase, repeat=2))))

survery_header= ['type',
                'name',
                'label::English',
                'hint::English',
                'guidance_hint::English',
                'label::Francais',
                'hint::Francais',
                'guidance_hint::Francais',
                'display_name',
                'choice_filter',
                'constraint',
                'constraint_message',
                'relevant',
                'repeat_count',
                'default',
                'readonly',
                'appearance',
                'parameters',
                'autoplay',
                'body::accuracyThreshold',
                'body::intent',
                'required',
                'required_message',
                'calculation',
                'media::image::English',
                'media::video::English',
                'media::audio::English',
                'media::image::Francais',
                'media::video::Francais',
                'media::audio::Francais'
                ]
choice_headers =[
    'list name',
    'name',
    'display_name',
    'label::English',
    'label::Francais',
    'media::image::English',
    'media::video::English',
    'media::audio::English',
    'media::image::Francais',
    'media::video::Francais',
    'media::audio::Francais'
]

settings_headers=[
    'list name',
    'name',
    'display_name',
    'label::English',
    'label::Francais',
    'media::image::English',
    'media::video::English',
    'media::audio::English',
    'media::image::Francais',
    'media::video::Francais',
    'media::audio::Francais'

]
survey_header_length = len(survery_header)
choice_header_length = len(choice_headers)
settings_header_length = len(settings_headers)



print(survey_header_length)
print(choice_header_length)
print(settings_header_length)



def create_survey(survey_name):
    workbook =xlsxwriter.Workbook(survey_name)
    # Add a bold format to use to highlight cells.
    bold =workbook.add_format({'bold':True})

    survey = workbook.add_worksheet('survey')
    choices = workbook.add_worksheet('choices')
    settings = workbook.add_worksheet('settings')
    # define headers for survey
    for i in range(survey_header_length):
        survey.write(excel_columns[i]+'1', survery_header[i],bold)

    

    workbook.close()

create_survey('sample.xlsx')
