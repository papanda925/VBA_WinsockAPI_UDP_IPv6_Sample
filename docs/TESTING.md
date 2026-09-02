# 動作確認手順

## 事前確認

- Windows版Excel 2010以降
- `.xlsm`または`.xlsb`で保存
- `src`の3モジュールをインポート
- VBAProjectのコンパイル成功
- 送受信先が`::1:60053`
- `CheckUdpIPv6StructureSizes`を実行し、`SOCKADDR_IN6=28 bytes`と表示された

## 正常系

1. `StartUdpIPv6ServerInNewExcel`を実行する。
2. 受信側Traceの`SERVER_READY`を確認する。
3. `SendHello`を実行し、応答を確認する。
4. `SendEchoSample`を実行する。
5. 日本語が文字化けせず返信されることを確認する。
6. 送信側・受信側Traceで同じTraceIdを探す。
7. `StopUdpIPv6Server`を実行する。
8. 受信側Traceの`SHUTDOWN`を確認する。
9. `CloseUdpIPv6ServerExcel`を実行し、専用Excelが終了することを確認する。

## UDPの特性を確認する試験

1. 受信側を停止した状態で`SendHello`を実行する。
2. `DATAGRAM_SENT`まで記録されることを確認する。
3. 約5秒後、応答待ちがタイムアウトすることを確認する。
4. 「送信API成功＝相手の受信成功ではない」ことを確認する。

## 確認結果

| 項目 | 環境 | 結果 | 備考 |
| --- | --- | --- | --- |
| VBAコンパイル | Office 32bit | 未確認 | 実機確認待ち |
| VBAコンパイル | Office 64bit | 未確認 | 実機確認待ち |
| HELLO | Windows / Excel | 未確認 | 実機確認待ち |
| 日本語ECHO | Windows / Excel | 未確認 | 実機確認待ち |
| 応答タイムアウト | Windows / Excel | 未確認 | 実機確認待ち |
| QUIT | Windows / Excel | 未確認 | 実機確認待ち |
