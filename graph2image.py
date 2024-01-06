import streamlit as st
import pandas as pd
import numpy as np
import plotly.figure_factory as ff

# ランダムな値が入ったデータフレームを作成
df = pd.DataFrame(
    np.random.rand(20, 3),
    columns=['a', 'b', 'c']
)

# データフレームを表示
st.line_chart(df)

# ランダムな値が入ったデータを生成
x1 = np.random.randn(200) - 2
x2 = np.random.randn(200)
x3 = np.random.randn(200) + 2
hist_data = [x1, x2, x3]
group_labels = ['Group 1', 'Group 2', 'Group 3']

# distplotでヒストグラムを描画する。Figure Factoryを使用してPlotlyのFigureを作成する
fig = ff.create_distplot(hist_data, group_labels, bin_size=[.1, .25, .5])

# Streamlitで作成したFigureを表示する
st.plotly_chart(fig, use_container_width=True)

# 作成したFigureを画像として保存する
fig.write_image("tmp/fig1.png")
