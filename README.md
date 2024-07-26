# 航空公園コート予約システム

航空公園のコートを予約作業を自動化するプログラムです。Windowsで使うことを想定しています。

## 機能
1. 他の人の投票数を一覧にする
2. 自動で抽選の投票をする
3. 自動で投票の確定作業をする

## 初期化の手順
1. リンクを参考に、コマンドプロンプトにnpmをインストール
   https://qiita.com/gahoh/items/8444da99a1f93b6493b4
3. コマンドプロンプトにgitをインストール(手順2まで)<br>
   https://qiita.com/T-H9703EnAc/items/4fbe6593d42f9a844b1c
4. コマンドプロンプトを開き、以下のコマンドを実行
   ```console
   git clone https://github.com/yukiobata1/res_Tokorozawa/
   cd res_Tokorozawa
   npm install
   ```
## 各機能の使い方
以下の作業はコマンドプロンプトを開いて、`cd res_Tokorozawa`としてから実行してください。
結果は、ローカルディスク>ユーザー>{ユーザー名}>res_Tokorozawaフォルダの中に出力されます。ピン止めしておくとわかりやすいです。<br>

### 他の人の投票数を一覧にする<br>
1. `node getDraft.mjs`を実行。
2. 30分程度で、`YYYY-MM`(2024-08など)というフォーマットのフォルダの中に'下書き.xlsx'が出力されます。
3. この中に投票数が書き込まれています。
### 自動で抽選の投票をする<br>
1. `YYYY-MM`のフォルダ内で、まず下書き.xlsxをコピーする
3. コピーされたファイルのセルを、投票する票数に書き換える
4. 投票先一覧.xlsxという名前に変更して、下書き.xlsxと同じフォルダ内に保存する
5. コマンドプロンプトに移動し、`nohup node vote.mjs &`を実行する
6. `YYYY-MM`フォルダの中に、投票の進行状況であるvoteRecord.txtが出力されます。各エントリの"done"が投票されたかに対応します
7. 8時間くらいで完了します。


注意点
- パソコンの電源を切ると処理が中断されてしまうので、スリープしない設定にしてつけっぱなしにしておいてください。
- 中断された際は、voteRecord.txtに実行状況が記録されているので、再度`nohup node vote.mjs &`を実行すればよいです。
- voteRecord.txtは手で編集しないでください。
- 投票数の合計は、アカウント数*4より少なくしてください。エラー制御を実装してないので何が起こるかわかりません。
- 期限は毎月9日です。メンテナンスの可能性があるので余裕をもってやった方がいいと思います。

### 投票の確認
1. コマンドプロンプトに移動し、`node confirm.mjs`を実行する
2. 2時間くらいで完了します。

### アカウント一覧の編集
- フォルダ内の、所沢アカウント一覧.xlsxに行を追加してください。
- また使いたくないアカウントがあればその行を削除して下さい。

わからないことがあればyukiobata1@gmail.comまでメールを送ってください。
