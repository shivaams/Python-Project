import base64
import uuid
import re

def download_button(object_to_download, download_filename, button_text):

    try:
        # some strings <-> bytes conversions necessary here
        b64 = base64.b64encode(object_to_download.encode()).decode()

    except AttributeError as e:
        b64 = base64.b64encode(object_to_download).decode()

    button_uuid = str(uuid.uuid4()).replace('-', '')
    button_id = re.sub('\d+', '', button_uuid)

    custom_css = f""" 
        <style>
            #{button_id} {{
                text-decoration: none;
                display: inline-flex;
                -webkit-box-align: center;
                align-items: center;
                -webkit-box-pack: center;
                justify-content: center;
                font-weight: 400;
                padding: 0.25rem 0.75rem;
                border-radius: 0.25rem;
                margin: 0px;
                line-height: 1.6;
                color: inherit;
                width: auto;
                user-select: none;
                background-color: rgb(19, 23, 32);
                border: 1px solid rgba(250, 250, 250, 0.2);
            }} 
            #{button_id}:hover {{
                border-color: #00ADB5;
                color: #00ADB5;
            }}
            #{button_id}:focus:not(:active) {{
                border-color: #00ADB5;
                color: #00ADB5;
            }}
            #{button_id}:focus {{
                box-shadow: #00ADB5 0px 0px 0px 0.2rem;
            }}
            #{button_id}:active {{
                box-shadow: none;
                background-color: #00ADB5;
                color: white;
                }}
        </style> """

    dl_link = custom_css + f'<a download="{download_filename}" id="{button_id}" href="data:file/txt;base64,{b64}">{button_text}</a><br></br>'

    return dl_link
