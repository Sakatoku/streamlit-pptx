from pptx import  Presentation
import streamlit as st
import pandas as pd
import numpy as np
import plotly.figure_factory as ff
from PIL import Image
# import plotly
# import kaleido

# このプロトタイプで使用する定数
OUTPUT_FILENAME = 'result.pptx'
OUTPUT_DIR = 'output'
TMP_DIR = 'tmp'
SLIDE_WIDTH = 16
SLIDE_HEIGHT = 9
EMU_PER_INCH = 914400
 # EMU_PER_CM = 360000

# PlotlyのFigureを作成する
def generate_plotly_figure(df):
    # ランダムな値が入ったデータを生成
    x1 = np.random.randn(200) - 2
    x2 = np.random.randn(200)
    x3 = np.random.randn(200) + 2
    hist_data = [x1, x2, x3]
    group_labels = ['Group 1', 'Group 2', 'Group 3']

    # distplotでヒストグラムを描画する。Figure Factoryを使用してPlotlyのFigureを作成する
    fig = ff.create_distplot(hist_data, group_labels, bin_size=[.1, .25, .5])
    return fig

# 画像を中央に貼り付けるためのサイズを計算する
def calculate_size(img, slide_size):
    # 画像の縦横比を計算する
    img_size = Image.open(img).size
    img_aspect_ratio = img_size[0] / img_size[1]

    # スライドのサイズを取得する
    slide_width, slide_height = slide_size
    # スライドの縦横比を計算する
    slide_aspect_ratio = slide_width / slide_height

    # 画像の縦横比とスライドの縦横比を比較する
    if img_aspect_ratio > slide_aspect_ratio:
        # 画像のほうが横長の場合、スライドの幅に合わせて画像を縮小する
        plot_size = (slide_width, slide_width / img_aspect_ratio)
    else:
        # 画像のほうが縦長の場合、スライドの高さに合わせて画像を縮小する
        plot_size = (slide_height * img_aspect_ratio, slide_height)

    # スライドのサイズと調整後の画像のサイズから、画像を中央に配置するための座標を計算する
    x = (slide_width - plot_size[0]) / 2
    y = (slide_height - plot_size[1]) / 2
    return x, y, plot_size[0], plot_size[1]

# 画像を貼り付けたプレゼンテーションを作成する
def make_presentation(imgs):
    # プレゼンテーションを新規作成する
    prs = Presentation()

    # スライドのサイズを設定する
    prs.slide_width = SLIDE_WIDTH  * EMU_PER_INCH
    prs.slide_height = SLIDE_HEIGHT * EMU_PER_INCH

    # レイアウトを選択。新規作成した場合は[6]が空スライド
    layout = prs.slide_layouts[6]

    # 画像ごとにスライドを作成する
    for img in imgs:
        slide = prs.slides.add_slide(layout)
        shapes = slide.shapes
        size = calculate_size(img, [prs.slide_width, prs.slide_height])
        st.write(img)
        st.write(size)
        _ = shapes.add_picture(img, size[0], size[1], size[2], size[3])
    
    return prs

# Figureを作成する
fig = generate_plotly_figure(None)
st.plotly_chart(fig, use_container_width=True)

# 作成したFigureを画像として保存する
st.write('start to save image')
fig.write_image(f'{TMP_DIR}/fig1.png', format='png', engine='orca')
st.write('finish to save image')

# 画像を中央に貼り付けたプレゼンテーションを作成する
prs = make_presentation([f'{TMP_DIR}/fig1.png'])

# プレゼンテーションファイルを保存する
prs.save(f'{OUTPUT_DIR}/{OUTPUT_FILENAME}')

# 保存したプレゼンテーションをダウンロードできるようにする
with open(f'{OUTPUT_DIR}/{OUTPUT_FILENAME}', 'rb') as f:
    bytes = f.read()
    st.download_button(
        label='Download PPTX',
        data=bytes,
        file_name=OUTPUT_FILENAME,
        mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
