'''
 * Copyright 2021 Ally Financial, Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 '''

import PySimpleGUI as sg
import base64

def convert_file_to_base64(filename):
    try:
        contents = open(filename, 'rb').read()
        encoded = base64.b64encode(contents)
        sg.clipboard_set(encoded)
        # pyperclip.copy(str(encoded))
        sg.popup('Copied to your clipboard!', 'Keep window open until you have pasted the base64 bytestring')
    except Exception as error:
        sg.popup_error('Cancelled - An error occurred', error)


if __name__ == '__main__':
    filename = sg.popup_get_file('Source Image will be encoded and results placed on clipboard', title='Base64 Encoder')

    if filename:
        convert_file_to_base64(filename)
    else:
        sg.popup_cancel('Cancelled - No valid file entered')
        
