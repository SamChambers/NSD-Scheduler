

import scheduler_functions as sf

import unittest


class TestSum(unittest.TestCase):

    def test_iphone_fix(self):

        input_text = '''Hello Sam Chambers,

A new schedule has been published for the week of Sun Jan 3, 2021. Your schedule is:

Mon Jan 4, 2021

7:00 AM - 12:30 PM - Morning

'''
        input_text_change = '''> Hello Sam Chambers,
>=20
> A new schedule has been published for the week of Sun Jan 3, 2021. Your sc='''
        input_text_change +='\r'
        input_text_change += '''
hedule is:
>=20
> Mon Jan 4, 2021
>=20
> 7:00 AM - 12:30 PM - Morning

'''

        output = sf.iphone_fix(input_text)

        self.assertEqual(output, input_text,
                         "Output should have been the same as the input")

        output = sf.iphone_fix(input_text_change)

        self.assertEqual(output, input_text,
                         "Output should have been changed to the new format correctly")
        

if __name__ == '__main__':
    unittest.main()
