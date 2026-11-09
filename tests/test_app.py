import io
import unittest
from app import app
from converter import parse_csv, parse_moodle_xml, export_data, ConversionError

CSV=b'name,type,question,correct,option_a,option_b\nQ1,single,Pilih jawaban,A,Ya,Tidak\n'
class ConverterTests(unittest.TestCase):
    def test_csv_xml(self):
        parsed=parse_csv(CSV)
        self.assertEqual(len(parsed['questions']),1)
        result,typ,ext=export_data(parsed['questions'],'xml','Demo')
        self.assertEqual(ext,'xml')
        self.assertEqual(len(parse_moodle_xml(result)['questions']),1)
    def test_incomplete_key_blocked(self):
        parsed=parse_csv(b'name,type,question,correct,option_a,option_b\nQ1,single,Pilih jawaban,,Ya,Tidak\n')
        with self.assertRaises(ConversionError): export_data(parsed['questions'],'xml','Demo')
    def test_http(self):
        client=app.test_client()
        self.assertEqual(client.get('/').status_code,200)
        self.assertEqual(client.get('/api/health').json['status'],'ok')
        r=client.post('/api/parse',data={'file':(io.BytesIO(CSV),'questions.csv')},content_type='multipart/form-data')
        self.assertEqual(r.status_code,200)
        self.assertEqual(len(r.json['questions']),1)
if __name__=='__main__': unittest.main()
