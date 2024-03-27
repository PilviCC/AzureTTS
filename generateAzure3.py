import pandas as pd
import azure.cognitiveservices.speech as speechsdk

#that need to be updated!!!
speech_key = "KEY"
service_region = "ServiceRegion"

speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)

excel_file = 'prompts.xlsx'

df = pd.read_excel(excel_file)

for index, row in df.iterrows():
    file_name = row['FileName']
    text = row['Text']

    audio_config = speechsdk.audio.AudioOutputConfig(filename=f"{file_name}.wav")

    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)

    speech_synthesis_result = speech_synthesizer.speak_text_async(text).get()

    if speech_synthesis_result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        print(f"Speech synthesized for text '{text}' and saved as '{file_name}.wav'")
    elif speech_synthesis_result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = speech_synthesis_result.cancellation_details
        print(f"Speech synthesis canceled: {cancellation_details.reason}")
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print(f"Error details: {cancellation_details.error_details}")
