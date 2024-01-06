from pptx import  Presentation

TEST_FILENAME = 'output/test.pptx'
prs = Presentation()

# レイアウトを確認する
layouts = prs.slide_layouts
if len(layouts) <= 0:
    print('No layout found')
    exit(1)

# 指定してスライドを追加する
prs.slides.add_slide(layouts[6])

# スライド内のシェイプにアクセスする
shapes = prs.slides[0].shapes

# まずはテンポラリフォルダに何かの画像が置かれている前提でプロトタイピングする
image_file = 'tmp/cat.png'
EMU_PER_CM = 360000
picture = shapes.add_picture(image_file, 10 * EMU_PER_CM, 6 * EMU_PER_CM, 15 * EMU_PER_CM, 8 * EMU_PER_CM)
# トリミングと回転を試す
picture.crop_bottom = 0.25
picture.rotation = -60

# まずはテンポラリフォルダに何かの画像が置かれている前提でプロトタイピングする
image_file = 'tmp/fig1.png'
picture = shapes.add_picture(image_file, 10 * EMU_PER_CM, 6 * EMU_PER_CM, 15 * EMU_PER_CM, 8 * EMU_PER_CM)

# プレゼンテーションファイルを保存する
prs.save(TEST_FILENAME)
