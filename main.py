from io import BytesIO
from docxtpl import DocxTemplate
from google.cloud import storage
import uuid


def download_template(template_filename):
    client = storage.Client()
    bucket = client.bucket('word-templates-bucket')
    blob = bucket.blob('requests/templates/' + template_filename)
    file_bytes = blob.download_as_bytes()
    file_io = BytesIO(file_bytes)

    return file_io


def get_npa(npa, lang):
    if lang == 'kg':
        d = {
            'n_p_a_1': 'Мамлекеттик органдардын жана жергиликтүү өз алдынча башкаруу органдарынын карамагындагы '
                       'маалыматтарга жетүү жөнүндө" Мыйзамдын 10-беренесинде белгиленген эки жумалык мөөнөт.',
            'n_p_a_2': 'Эл аралык стандарттарга ылайык, ыйгарым укуктуу мамлекеттик орган суралып жаткан маалыматтар '
                       'ачык деген негизде маалымат берүүдөн баш тартууга укугу жок.',
            'n_p_a_3': 'Мамлекеттик органдардын жана жергиликтүү өз алдынча башкаруу органдарынын карамагындагы '
                       'маалыматтарга жетүү жөнүндө.',
            'n_p_a_4': 'Суралган маалыматты Кыргыз Республикасынын 05.12.1997-жылдагы "аалыматка жетүүнүн '
                       'кепилдиктери жана эркиндиги жөнүндө" мыйзамына ылайык бериңиз.'
        }
    else:
        d = {
            'n_p_a_1': 'Запрашиваемую информацию просим предоставить в двухнедельный срок установленный ст. 10 закона '
                       '«О доступе к информации находящейся в ведении государственных органов и органов МСУ».',
            'n_p_a_2': 'Обращаем Ваше внимание, что согласно международным нормам, уполномоченное государственное '
                       'учреждение не вправе отказывать в предоставлении информации на основании того, '
                       'что запрашиваемые данные находятся в открытом доступе, то есть, уже были опубликованы '
                       'где-либо.',
            'n_p_a_3': 'Напоминаем, что согласно ст.10 закона «О доступе к информации находящейся в ведении '
                       'государственных органов и органов МСУ», ответ на письменный запрос о предоставлении '
                       'информации должен носить исчерпывающий характер, исключающий необходимость повторного '
                       'обращения заинтересованного лица по тому же предмету запроса.',
            'n_p_a_4': 'Запрашиваемую информацию просим предоставить согласно Закону Кыргызской Республики «О '
                       'гарантиях и свободе доступа к информации» от 05.12.1997 года.'
        }
    return '\n\t' + '\n\t'.join(d[item] for item in npa)


def generate_word_doc(request_json):
    if request_json['which_language'] == 'kyrgyz':
        doc = DocxTemplate(download_template('template_kg'))
        npa = get_npa(request_json['n_p_a'], 'kg')
        context = {
            'number': request_json['number'],
            'address': request_json['address'],
            'date': request_json['date'],
            'body': request_json['body'],
            'npa': npa

        }
        doc.render(context)
    else:
        doc = DocxTemplate(download_template('template_ru'))
        npa = get_npa(request_json['n_p_a'], 'ru')
        context = {
            'number': request_json['number'],
            'address': request_json['address'],
            'date': request_json['date'],
            'body': request_json['body'],
            'npa': npa
        }
        doc.render(context)
    return doc


def main(request):
    response = {}
    request_json = request.get_json()[0] if isinstance(request.get_json(), list) else request.get_json()

    if request_json:
        print(request_json)

    document = generate_word_doc(request_json)

    save_file_stream = BytesIO()
    # Save the .docx to the buffer
    document.save(save_file_stream)
    # Reset the buffer's file-pointer to the beginning of the file
    save_file_stream.seek(0)

    client = storage.Client()
    bucket = client.bucket('word-templates-bucket')

    doc_uuid = str(uuid.uuid4())

    blob_save = bucket.blob('requests/results/' + doc_uuid + '.docx')
    blob_save.upload_from_string(save_file_stream.getvalue())
    blob_save.make_public()
    url = blob_save.public_url

    response['docx_link'] = {'file': 'requests/results/' + doc_uuid + '.docx'}

    return {'file': url}
