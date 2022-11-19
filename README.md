# ProgressBarAddins
PowerPointやExcelのオートシェイプで塗りつぶしによって進捗を表現する単機能的なアドインです。

<img src="https://user-images.githubusercontent.com/46097651/202833209-31f6a83d-1c0b-4d46-9264-4eb950a4dccf.png" width="400px">

## 使用方法
アドインをインストールすると シェイプを選択した時に出てくる「図の形式」タブに「進捗でFill」ボタンが表示されるようになります。

![image](https://user-images.githubusercontent.com/46097651/202834038-4cd8766e-246d-46cc-be70-3963b97a515e.png)

## 仕組み
グラデーション設定によって2色の色を設定しています。
Officeの書式設定から行うと、1%刻みでしか設定できないため、境界がどうしてもぼやけてしまいます。
(同じ%を設定することで上手く動作することもあります)


プログラム内部的には小数でも設定できますので、本アドインでは小数点刻みでグラデーションの境界を設定することで2色の塗りつぶしを表現しています。

<img src="https://user-images.githubusercontent.com/46097651/202833650-f73f2fcc-3f21-4483-834f-f0f9163f6676.png" width="250px">
2色の塗りつぶしの色は元の図形の枠線色と塗りつぶし色を使用します。
<img src="https://user-images.githubusercontent.com/46097651/202833829-70e36c9e-c7e4-49ef-966f-d57a5194fe4b.png" width="300px">
