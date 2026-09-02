# Excel VBA Winsock UDP/IPv6 Client–Server Sample

Excel VBAからWindowsのWinsock APIを直接呼び出し、IPv6のUDP送信側と簡易サーバーをローカルPC内で動かす学習用サンプルです。

`WScript.Shell`、PowerShell、`curl`、外部ActiveXコントロールは使用しません。送受信の各段階を`Trace`シートで確認でき、UDPの「送信成功と受信成功は別」という性質も実際に試せます。

> [!IMPORTANT]
> UDPは到達、順序、重複排除、再送を保証しません。このサンプルは学習用で、認証や暗号化もありません。初めはIPv6ループバックアドレス`::1`から変更しないでください。

## 学べること

- IPv6ループバックアドレス`::1`
- `SOCKADDR_IN6`と16バイトのIPv6アドレス
- `socket(AF_INET6, SOCK_DGRAM, IPPROTO_UDP)`
- `bind` / `sendto` / `recvfrom` / `closesocket`
- UDPには`listen`、`accept`、通常の接続確立がないこと
- 送信元IPv6アドレスとポートの取得
- UTF-16とUTF-8の変換
- 応答待ちタイムアウトの必要性
- TraceIdによる送信側・受信側ログの照合

## ファイル構成

| ファイル | 内容 |
| --- | --- |
| `src/VBA_WinsockAPI_UDP_IPv6_Sample.bas` | IPv6 UDPサーバー、クライアント、実行マクロ |
| `src/TraceLogger.bas` | Traceシートとイミディエイトウィンドウへの記録 |
| `src/Utf8Codec.bas` | VBA文字列とUTF-8バイト列の相互変換 |
| `docs/TESTING.md` | 動作確認項目と期待結果 |
| `docs/REVIEW.md` | 10ペルソナ・100観点レビューの記録 |
| `SECURITY.md` | 安全上の注意と脆弱性の連絡方法 |

## 動作環境

- Windows 10またはWindows 11
- Excel 2010以降（VBA 7）
- 32ビット版または64ビット版Office
- IPv6が有効なWindows環境
- `.xlsm`または`.xlsb`形式のマクロ有効ブック

macOS版Excelでは動作しません。

## 使い方

### 1. モジュールを取り込む

1. マクロ有効ブックを作成し、一度保存します。
2. `Alt` + `F11`でVisual Basic Editorを開きます。
3. 「ファイル」→「ファイルのインポート」から`src`内の3ファイルを取り込みます。
4. 「デバッグ」→「VBAProjectのコンパイル」を実行します。

GitHub上の`.bas`はUTF-8です。VBEで日本語が文字化けする場合は、UTF-8対応エディターでCP932（Shift_JIS）へ変換してからインポートするか、VBEの新規モジュールへ貼り付けてください。TCP版とUDP版の同名マクロを衝突させないため、両リポジトリは別々のブックへ取り込みます。

### 2. 受信側を起動する

`StartUdpIPv6ServerInNewExcel`を実行します。

同じブックが別のExcelプロセスで読み取り専用として開き、1秒後に`[::1]:60053`でデータグラムを待ちます。受信側Excelの`Trace`シートに`SERVER_READY`が表示されるまで待ってください。

### 3. データグラムを送る

| マクロ | 内容 | 受信側の応答 |
| --- | --- | --- |
| `SendHello` | 到達確認 | 固定HELLOメッセージ |
| `SendEchoSample` | 日本語をUTF-8で送信 | 受信内容をそのまま返信 |
| `StopUdpIPv6Server` | 終了要求 | 応答後に受信ループ終了 |
| `CloseUdpIPv6ServerExcel` | 別Excelの終了 | 読み取り専用ブックを保存せず閉じる |

送信側は`sendto`の後、受信側からの応答を最大5秒待ちます。受信側を起動せずに送ると、`sendto`自体は成功しても応答待ちがタイムアウトします。これがUDPの重要な学習ポイントです。

終了時は`StopUdpIPv6Server`を実行し、受信側Traceの`SHUTDOWN`を確認してから`CloseUdpIPv6ServerExcel`を実行します。

## 通信の流れ

### 受信側

```text
WSAStartup
  → socket(AF_INET6, SOCK_DGRAM)
  → IPV6_V6ONLY
  → bind([::1]:60053)
  → recvfrom
  → sendto
  → closesocket
  → WSACleanup
```

### 送信側

```text
WSAStartup
  → socket(AF_INET6, SOCK_DGRAM)
  → sendto([::1]:60053)
  → recvfrom（応答確認）
  → closesocket
  → WSACleanup
```

## TCP版との違い

| 項目 | UDP | TCP |
| --- | --- | --- |
| 接続確立 | なし | `connect` / `accept` |
| 到達保証 | なし | 再送などをTCPが担当 |
| 順序保証 | なし | あり |
| データ境界 | データグラム単位 | バイトストリーム |
| 主なAPI | `sendto` / `recvfrom` | `send` / `recv` |

TCP/IPv6版は[`VBA_WinsockAPI_TCP_IPv6_Sample`](https://github.com/papanda925/VBA_WinsockAPI_TCP_IPv6_Sample)にあります。

## Traceシート

`Time`、`TraceId`、`Side`、`Step`、`Direction`、`Detail`、`Result`、`ElapsedMs`を記録します。要求電文にTraceIdを含めるため、送信側と受信側で同じデータグラムを探せます。

代表的なStepは次のとおりです。

```text
CLIENT: STARTUP → DATAGRAM_SENT → DATAGRAM_RECEIVED
SERVER: SERVER_READY → DATAGRAM_RECEIVED → DATAGRAM_SENT → SHUTDOWN
```

## 制限事項

- 1データグラムは8,192バイト以下を想定
- IPv6フラグメンテーションやPath MTU Discoveryの検証なし
- 再送、順序制御、重複排除なし
- 複数要求の並行処理なし
- 認証・暗号化なし
- 受信側の`recvfrom`は同期・ブロッキング処理

## トラブルシューティング

### 送信したが応答がない

受信側Traceに`SERVER_READY`があるか、両方のポート番号が`60053`か確認してください。UDPでは受信側が起動する前のデータグラムは保持されません。

### Winsockエラー10060

応答待ちのタイムアウトです。受信側が起動していない、アドレス・ポートが違う、または応答が遮断された可能性があります。

### ポートを使用中というエラー

以前の受信側Excelが残っていないか確認します。他のアプリが`60053`を使用している場合は、送信側と受信側で同じ未使用ポートへ変更してください。

### 別Excelでマクロを実行できない

組織ポリシーやトラストセンター設定により、自動で開いた読み取り専用ブックのマクロが禁止される場合があります。設定を回避せず、信頼できる場所と管理者の方針を確認してください。

## 参考資料

- [IPv6 guide for Windows Sockets applications](https://learn.microsoft.com/windows/win32/winsock/ip-version-6-2)
- [SOCKADDR_IN6 structure](https://learn.microsoft.com/windows/win32/api/ws2ipdef/ns-ws2ipdef-sockaddr_in6_lh)
- [sendto function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-sendto)
- [recvfrom function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-recvfrom)
- [InetPtonW function](https://learn.microsoft.com/windows/win32/api/ws2tcpip/nf-ws2tcpip-inetptonw)
- [64-bit VBA overview](https://learn.microsoft.com/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview)

## 実機確認について

コードとAPI宣言は静的に確認していますが、この作成環境にはWindows版Excelがありません。`docs/TESTING.md`に従い、実機でコンパイルとループバック通信を確認してください。
