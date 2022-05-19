from multiprocessing.connection import wait
import aspose.slides as slides
import pandas as pd
import requests
from flask import Flask, request, Blueprint, jsonify
import json
from flask_cors import CORS
import time

app = Flask(__name__)
CORS(app)


@app.route('/', methods=['POST', 'GET'])
def main():
    if request.method == 'POST':
        output = request.get_json()
        output = {
            'excel_file': output['y'], 'ppt_file': output['x'], 'credential': output['z']}

        # x = {"response": botnoi_voice(
        #     excel_file, ppt_file, credential)['audio_url']}
        scripts_df = format_scripts_file(
            output['excel_file'], output['credential'])
        for url, filename in zip(scripts_df['audio_url'].values, scripts_df['voice_file_name'].values):
            DownloadFile(url, f'./voices/{filename}')
        voiceNames = format_voice_name(
            scripts_df['slide'], scripts_df['voice_file_name'])
        result = embed_voice_in_pptx(output['ppt_file'], voiceNames)

        return result
    if request.method == 'GET':
        q = request.args.get('q')
        return jsonify({'result': q})


def botnoi_voice(sentence, speaker, credential):
    url = "https://voice.botnoi.ai/api/service/generate_audio"
    payload = {"text": sentence, "speaker": speaker,
               "volume": 1, "speed": 1, "type_media": "wav"}
    headers = {
        'Botnoi-Token': credential
    }
    response = requests.request("POST", url, headers=headers, json=payload)
    return response.json()


def format_scripts_file(filepath, credential):
    speakerEnum = {
        'คริส': 27,
        'ครูดีดี๊': 17,
        'คอรีย์': 57,
        'คุณงาม': 3,
        'นายเบรด': 35,
        'นาเดียร์': 10,
        'น้าเกรซ': 31,
        'ปีเตอร์': 21,
        'ผู้ใหญ่ลี': 40,
        'วนิลา': 12,
        'สาลี': 37,
        'สโม๊ค': 33,
        'หญิงไอโกะ': 30,
        'อนันดา': 14,
        'อลัน': 5,
        'อลิสา': 8,
        'อาจารย์หลิน': 44,
        'อาวอร์ม': 20,
        'ฮิโระ': 16,
        'เจสัน': 29,
        'เจ้าเนิร์ด': 18,
        'เท็ดดี้': 41,
        'เนโอ': 36,
        'เบลล์': 59,
        'เลโอ': 9,
        'เอวา': 1,
        'แบมบู': 38,
        'แม็กซ์': 4,
        'โตโต้': 42,
        'โนรา': 43,
        'โบ': 2,
        'โอโตะ': 19,
        'ไซเรน': 6,
        'ไอลีน': 15
    }
    scripts = pd.read_excel(filepath)
    scripts['slide'] = scripts['slide'].values.astype('int64')
    scripts['audio_url'] = [botnoi_voice(sentence, speakerEnum[speaker], credential)['audio_url'] for sentence, speaker in zip(
        scripts['sentence'].values, scripts['speaker'].values)]
    scripts['voice_file_name'] = scripts['audio_url'].apply(
        lambda url: url.split('_')[-1])
    return scripts


def format_voice_name(slideSeries, voiceFileNameSeries):
    # pd.Series([1, 1, 2, 2, 3]), pd.Series(['05012022080420464300.wav', '05012022080423553980.wav',
    #                                                                '05012022080429180110.wav', '05012022080432791033.wav', '05012022080435933073.wav'])
    result = []
    index = 0

    for sld in list(slideSeries.value_counts().sort_index().index):
        slide = []
        for voiceName in range(list(slideSeries.value_counts().sort_index().values)[sld-1]):
            slide.append(list(voiceFileNameSeries.values)[index])
            index += 1
        result.append(slide)

    return result


def DownloadFile(url, fn):
    # local_filename = 'test.mp3'
    r = requests.get(url)
    with open(fn, 'wb') as f:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk:  # filter out keep-alive new chunks
                f.write(chunk)


def embed_voice_in_pptx(filepath, voiceName2DArray):
    print(voiceName2DArray)
    # load presentation
    with slides.Presentation(filepath) as presentation:
        try:
            # for slide in presentation.slides:
            for i in range(len(voiceName2DArray)):
                slide = presentation.slides[i]

                # load the wav sound file to stream
                x_axis = 50
                for voice in voiceName2DArray[i]:
                    with open(f'./voices/{voice}', "rb") as in_file:
                        # add audio frame
                        audio_frame = slide.shapes.add_audio_frame_embedded(
                            x_axis, 450, 30, 30, in_file)
                    x_axis += 50

                    # set play mode and volume of the audio
                    audio_frame.play_mode = slides.AudioPlayModePreset.AUTO
                    audio_frame.volume = slides.AudioVolumeMode.LOUD
        except:
            pass

        # write the PPTX file to disk
        presentation.save(filepath.split('.')[0]+'-embeded.pptx',
                          slides.export.SaveFormat.PPTX)
        return jsonify({'result': 'completed'})


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1")
