import pandas as pd
from openpyxl import load_workbook
from twilio.rest import Client
from dotenv import load_dotenv
import os
import time
from flask import Flask, request, jsonify, send_file
from twilio.twiml.voice_response import VoiceResponse
import threading
from pyngrok import ngrok
from pathlib import Path
import openai
import requests
from datetime import datetime
import shutil
from pydub import AudioSegment
from queue import Queue
import re
import ast

# 電話番号の冒頭の0を+81に変換する関数
def convert_phone_number(phone_number):
    print(f"電話番号を変換します: {phone_number}")
    phone_number = str(phone_number)  # 文字列に変換
    if phone_number.startswith('0'):
        converted = '+81' + phone_number[1:]
        print(f"変換後の電話番号: {converted}")
        return converted
    print("電話番号は変換の必要がありません")
    return phone_number

# 会話履歴から空き部屋の部屋番号を抽出する関数
def extract_vacant_rooms_from_conversation(conversation_history):
    """
    conversation_history.txtから会話履歴を読み込み、
    ChatGPTで分析して空き部屋の部屋番号をリストとして抽出する関数。
    
    :param file_path: conversation_history.txt のファイルパス
    :return: 空き部屋の部屋番号をリストとして返す
    """
    try:
        # ChatGPT APIを利用して、空き部屋の部屋番号を抽出するプロンプトを作成
        prompt = f"以下の会話履歴から、空き部屋として紹介されている部屋番号をすべてリストアップしてください。\n\n{conversation_history}"

        # ChatGPT APIを使用して会話の内容を分析
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role":"system","content":"あなたは不動産仲介会社の事務員です。"},
                {"role":"user","content":f'{prompt}'}
            ]
        )

        # ChatGPTの応答から空き部屋の部屋番号をリストとして抽出
        vacant_rooms_text = response.choices[0].message.content

        # 空き部屋の部屋番号をリストに変換
        vacant_rooms_list = [room.strip() for room in vacant_rooms_text.split('\n') if room]

        return vacant_rooms_list

    except Exception as e:
        print(f"会話履歴からの部屋番号取得時にエラーが発生しました: {e}")
        return []

# 会話履歴から空き部屋の各情報（結果、部屋番号）が十分に収集されているか判定する関数
def judgement_enough(file_path, conversation_history,model="gpt-4o-mini"):
    element="部屋番号"
    # 会話履歴をテキストファイルに保存
    save_conversation_history(file_path, conversation_history)
    # move_file_to_date_folder(conversation_history_file_path, f"{company_name}_{building_name}")

    # 会話履歴を読み込む
    with open(file_path, 'r', encoding='utf-8') as file:
        conversation_history = file.read()
    print(conversation_history)
    
    if element == "部屋番号":

        response = openai.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": f"あなたは不動産の専門家です。"},
                {"role": "user", "content": f'以下の会話履歴から、{element}に関する情報が十分に収集されているかどうかを判定し、「十分」または「不十分」の結果のみを出力してください。このとき、システムの、「〇〇号室の他に空いている部屋は無いか」といった旨の質問に対して否定的な回答を得た際には「十分」としてよい。\n{conversation_history}'}
            ]
        )
    else:
        response = openai.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": f"あなたは不動産の専門家です。"},
                {"role": "user", "content": f'以下の会話履歴から、{element}に関する情報が十分に収集されているかどうかを判定し、「十分」または「不十分」の結果のみを出力してください。\n{conversation_history}'}
            ]
        )
    print(response.choices[0].message.content)
    analysis_result = response.choices[0].message.content

    # "不十分" や "不足している" のような否定的な表現が含まれていないかも確認する
    if any(keyword in analysis_result for keyword in ["十分", "complete", "全ての情報", "sufficient"]) and \
       not any(neg in analysis_result for neg in ["不十分", "不足", "足りない", "missing", "incomplete"]):
        return True
    else:
        return False
  
# 音声ファイルを結合、またはrecording1が空の場合はrecording2のコピーを作成する関数
def concatenate_or_copy_recordings(recording1, recording2, output_file_path):
    """
    2つの録音ファイルを結合、またはrecording2が空の場合はrecording1のコピーを作成する関数
    
    :param recording1: 1つ目の録音ファイルのパス（空の場合はrecording1のコピーを作成）
    :param recording2: 2つ目の録音ファイルのパス
    :param output_file_path: 出力するファイルのパス
    """
    try:
        if not recording1:  # recording2が空の場合
            shutil.copyfile(recording2, output_file_path)
            print(f"{recording2}のコピーが作成されました: {output_file_path}")
        elif not recording2:  # recording1が空の場合
            shutil.copyfile(recording1, output_file_path)
            print(f"{recording1}のコピーが作成されました: {output_file_path}")
        elif recording1==output_file_path:
            output_file_path_backup=output_file_path.split(".")[0]+"_backup.mp3"
            shutil.copyfile(output_file_path, output_file_path_backup)
            # 各音声ファイルをロード
            audio1 = AudioSegment.from_file(output_file_path_backup)
            audio2 = AudioSegment.from_file(recording2)
            
            # 音声ファイルを結合
            combined_audio = audio1 + audio2
            
            # 結合した音声ファイルを保存
            combined_audio.export(output_file_path, format="mp3")
            print(f"音声ファイルが結合され、保存されました: {output_file_path}")
        else:
            # 各音声ファイルをロード
            audio1 = AudioSegment.from_file(recording1)
            audio2 = AudioSegment.from_file(recording2)
            
            # 音声ファイルを結合
            combined_audio = audio1 + audio2
            
            # 結合した音声ファイルを保存
            combined_audio.export(output_file_path, format="mp3")
            print(f"音声ファイルが結合され、保存されました: {output_file_path}")
    except Exception as e:
        print(f"音声ファイルの処理中にエラーが発生しました: {e}")

# 音声をテキスト化する関数
def transcribe_audio(audio_file_path: str, model: str = "whisper-1") -> str:
    print(f"音声ファイルをテキスト化しています: {audio_file_path}")
    with open(audio_file_path, "rb") as audio_file:
        transcript = openai.audio.transcriptions.create(
            model=model,
            file=audio_file,
            language="ja"  # 言語を日本語に指定
        )
    print("音声ファイルのテキスト化が完了しました")
    return transcript.text

# ChatGPT 4o miniを使用して応答を生成する関数
def generate_response(stance: str, conversation_history: list, model: str = "gpt-4o-mini") -> str:
    print("ChatGPT-4.0-miniを使用して応答を生成しています")
    response = openai.chat.completions.create(
      model=model,
      messages=[
        {"role": "system", "content": "あなたは不動産仲介会社の営業マンで、不動産管理会社に対して、管理会社が管理する物件の各部屋について空き状況を確認する目的で電話をかけています。"},
        *conversation_history,
        {"role": "user", "content": "話し相手は電話口の不動産仲介会社です。相手の発言に対して丁寧に、極力短い文字数で対応してください。あなたが話す文言だけを生成してくれればよいです。"+stance},
      ]
    )
    print("応答が生成されました")
    return response.choices[0].message.content

# テキストを音声ファイル化する関数
def text_to_speech(text: str, output_file_path: str, model: str = "tts-1", voice: str = "alloy"):
    print(f"テキストを音声ファイルに変換しています: {text}")
    speech_file_path = Path(output_file_path)
    
    # OpenAI APIを使用して音声を生成
    response = openai.audio.speech.create(
        model=model,
        voice=voice,
        input=text
    )
    with open(speech_file_path, 'wb') as f:
        f.write(response.content)
        print(f"音声ファイルが生成されました: {speech_file_path}")

# 会話履歴をテキストファイルに保存する関数
def save_conversation_history(file_path: str, conversation_history: list):
    try:
        print(f"会話履歴をファイルに保存します: {file_path}")
        with open(file_path, 'w', encoding='utf-8') as file:
            for entry in conversation_history:
                # roleとcontentのペアを取り出して、ファイルに書き込む
                role = entry.get('role', 'unknown')  # roleがない場合は'unknown'を使用
                content = entry.get('content', '')  # contentがない場合は空文字列を使用
                file.write(f"{role}: {content}\n")
        print(f"会話履歴が {file_path} に保存されました")
    except Exception as e:
        print(f"会話履歴の保存中にエラーが発生しました: {e}")

# 空のビルデータを格納したリストに空き部屋情報を追加する関数
def add_vacant_rooms_to_building(building_data,building_name, vacant_room_numbers):
    """
    vacant_rooms から部屋番号を抽出し、building_name ごとに対応する辞書を building_data リストに格納する関数
    
    :param building_name: ビル名
    :param vacant_rooms: 部屋番号が記載されたリスト
    :return: 更新された building_data
    """
    print(f"[INFO] Building name: {building_name}")
    print(f"[INFO] Vacant rooms before processing: {vacant_room_numbers}")

    # 部屋番号ごとの情報を作成
    room_info_list = [{'部屋番号': room, '坪数': '', '坪単価': '', '保証金': '', '入居時期': '', '備考': ''} for room in vacant_room_numbers]
    
    # building_data に辞書を追加
    building_data.append({building_name: room_info_list})
    
    return building_data

# ビルデータを格納したリスト内の辞書を、部屋番号数字昇順→アルファベット昇順でソートする関数
def sort_vacant_rooms_in_building(building_data):
    """
    building_data の中の各ビルの部屋番号を数字昇順→アルファベットの昇順でソートし、
    その辞書の順番を入れ替える関数。
    
    :param building_data: 部屋情報を含むビルごとの辞書のリスト
    :return: ソート後の building_data
    """
    
    def room_sort_key(room):
        # 部屋番号の数字部分とアルファベット部分を抽出してタプルを作成
        room_number = room['部屋番号']
        number_part = ''.join(filter(str.isdigit, room_number))
        alpha_part = ''.join(filter(str.isalpha, room_number))
        return (int(number_part) if number_part else 0, alpha_part)
    
    sorted_building_data = []
    
    for building in building_data:
        for building_name, rooms in building.items():
            # 部屋番号をソート
            sorted_rooms = sorted(rooms, key=room_sort_key)
            sorted_building_data.append({building_name: sorted_rooms})
    
    return sorted_building_data

# 指定文言で始まる全ての列を削除する関数（いまは空室情報固定）
def remove_vacant_columns(df):
    """
    DataFrameのカラム名が「空室情報」で始まる全ての列を削除する関数

    Parameters:
    - df: 操作対象のDataFrame

    Returns:
    - 更新後のDataFrame
    """
    # カラム名が「空室情報」で始まる列を見つけて削除
    cols_to_remove = [col for col in df.columns if col.startswith("空室情報")]
    print(f"削除対象のカラム: {cols_to_remove}")
    df.drop(columns=cols_to_remove, inplace=True)
    print("指定されたカラムを削除しました。")

    return df

# ビルデータ内情報をエクセルに書き込む関数
def insert_vacant_info_to_columns(building_data, df):
    """
    building_dataを元に、ビル名称が一致する行に対してH列から順に空室情報を挿入する関数

    Parameters:
    - building_data: 空室情報が格納されたリスト（辞書のリスト）
    - df: 操作対象のDataFrame

    Returns:
    - 更新後のDataFrame
    """
    for building in building_data:
        for building_name, room_info_list in building.items():
            # 「ビル名称」が一致する行を探す
            matching_rows = df[df['ビル名称'] == building_name]
            
            for index, row in matching_rows.iterrows():
                print(f"行 {index+1}: ビル名称は {building_name} です。")
                
                # H列から順に空室情報を格納していく
                for i, room_info in enumerate(room_info_list):
                    col_index = 7 + i  # H列は8番目のカラムでindexは7、I列は9番目のカラムでindexは8
                    col_name = df.columns[col_index] if col_index < len(df.columns) else None
                    
                    # 列が存在しない、またはカラム名が「空室情報」で始まっていない場合
                    if col_name is None or not col_name.startswith("空室情報"):
                        new_col_name = f"空室情報{i+1}"
                        df.insert(col_index, new_col_name, "")
                        col_name = new_col_name
                        print(f"新しいカラム '{col_name}' を挿入しました。")
                    
                    # 辞書内の情報をテキストにまとめる
                    room_info_text = "\n".join([f"{key}: {value}" for key, value in room_info.items()])
                    
                    # 指定列に情報を格納
                    df.at[index, col_name] = room_info_text
                    print(f"行 {index+1}, 列 {col_name} を更新しました: {room_info_text}")
    
    return df

# 会話履歴から部屋番号を抽出する関数
def extract_room_numbers(conversation_history):

    print(f"[INFO] Extracting room numbers from: {conversation_history}")

    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "あなたは、Pythonのリスト形式でデータを出力するAIです。"},
                {"role": "user", "content": f"以下の文に対して、Pythonのリスト形式で部屋番号を出力してください。リストは部屋番号のみを含み、引用符で囲まれた文字列形式で出力してください。\n例: ['101', '102', '103']\n{conversation_history}"}
            ]
        )
        print(f"[INFO] GPT-4 response received: {response}")
    except Exception as e:
        print(f"[ERROR] Failed to extract room numbers: {e}")
        return []

    # GPT-4からの応答をリスト形式で出力
    try:
        # 応答から部屋番号のリストを抽出
        content = response.choices[0].message.content
        print(f"[INFO] Extracted content: {content}")
        
        # 正規表現を使用してリストを抽出
        match = re.search(r'\[([^\[\]]+)\]', content)
        if match:
            room_numbers = match.group(1).split(',')
            room_numbers = [num.strip().strip('"').strip("'") for num in room_numbers]
            print(f"[INFO] Extracted room numbers: {room_numbers}")
            return room_numbers
        else:
            print("[ERROR] Failed to parse room numbers from GPT-4 response")
            return []
    except Exception as e:
        print(f"[ERROR] Failed to process GPT-4 response: {e}")
        return []

# 会話履歴から部屋の詳細情報を抽出し、辞書を生成する関数
def extract_room_details(missing_info,conversation_history):

    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "あなたは、Pythonの辞書形式でデータを出力するAIです。"},
                {"role": "user", "content": f"以下の文に対して、Pythonの辞書形式で{missing_info}のそれぞれのキーに対応する値を抽出し、正しい位置に格納してください。各値は引用符で囲まれた文字列形式で格納してください。\n{conversation_history}"}
            ]
        )
        print(f"[INFO] GPT-4 response received: {response}")
    except Exception as e:
        print(f"[ERROR] Failed to extract room numbers: {e}")
        return []

    # GPT-4からの応答をリスト形式で出力
    try:
        # 応答から部屋番号のリストを抽出
        content = response.choices[0].message.content
        print(f"[INFO] Extracted content: {content}")
        return content
        
    except Exception as e:
        print(f"[ERROR] Failed to process GPT-4 response: {e}")
        return []

# ファイル名を変えて、今日の日付のフォルダに移動フォルダに格納する関数
def move_file_to_date_folder(company_name,filepath, new_name):
    global room_numbers
    # 日付をYYYYMMDD形式で取得
    date_str = datetime.now().strftime('%Y%m%d')

    # 新しいフォルダパスを作成
    directory = os.path.dirname(filepath)
    date_folder = os.path.join(directory, date_str)
    company_folder=os.path.join(date_folder,company_name)

    # フォルダが存在しない場合は作成
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"フォルダを作成しました: {date_folder}")
    else:
        print(f"フォルダが既に存在します: {date_folder}")
    if not os.path.exists(company_folder):
        os.makedirs(company_folder)
        print(f"フォルダを作成しました: {company_folder}")
    else:
        print(f"フォルダが既に存在します: {company_folder}")

    # 元のファイルの拡張子を取得
    _, file_extension = os.path.splitext(filepath)

    # 新しいファイル名と拡張子を組み合わせる
    new_filename = new_name + file_extension
    new_filepath = os.path.join(company_folder, new_filename)

    # 同じ名前のファイルが存在する場合、(1), (2), ... のように番号を付ける
    counter = 1
    while os.path.exists(new_filepath):
        new_filename_with_suffix = f"{new_name}({counter}){file_extension}"
        new_filepath = os.path.join(company_folder, new_filename_with_suffix)
        counter += 1

    # ファイルを新しい場所に移動
    os.rename(filepath, new_filepath)
    print(f"ファイルを {filepath} から {new_filepath} に移動しました。")

    return new_filepath

# {}を含む文字列から辞書を生成する関数
def dict_from_str(room_detail):
    # 正規表現を使用して、{}で囲まれた部分を抽出
    dict_str_match = re.search(r'\{.*\}', room_detail, re.DOTALL)

    if dict_str_match:
        dict_str = dict_str_match.group()

        # インデントと不要な改行を削除
        dict_str = re.sub(r'^[ \t]+', '', dict_str, flags=re.MULTILINE)

        # シングルクォートをダブルクォートに変換 (JSON風にするため)
        dict_str = dict_str.replace("'", '"')

        # Noneをnullに変換 (JSON形式に準拠)
        dict_str = dict_str.replace('None', 'null')

        try:
            # 辞書として変換
            room_detail_dict = ast.literal_eval(dict_str)
            return room_detail_dict
        except Exception as e:
            print(f"辞書への変換に失敗しました: {e}")
            return None
    else:
        print("辞書が見つかりませんでした")
        return None
    
