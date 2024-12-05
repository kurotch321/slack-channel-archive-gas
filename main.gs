const sheetsApi = SpreadsheetApp.getActiveSpreadsheet();
const logSheet = sheetsApi.getSheetByName("ログ");
const targetSheet = sheetsApi.getSheetByName("対象チャンネル");
const channelSheet = sheetsApi.getSheetByName("チャンネル一覧");
const slackToken =
  PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");

function writeLog(execId, status, targetChannel, channelName, response) {
  const statusCode = status === true ? "完了" : "エラー"; // statusがtrueなら完了、falseならエラーにする
  return logSheet.appendRow([
    execId,
    statusCode,
    targetChannel,
    channelName,
    response,
  ]); // A列に実行ID、B列にステータス、C列にチャンネルID、D列にチャンネル名、E列にレスポンスを書き込む
}

function updateChannelButton() {
  const conversationsApi = callSlackApi(
    "conversations.list?exclude_archived=true&limit=1000"
  );
  const channels = conversationsApi.channels; // チャンネルを取り出す
  Logger.log("チャンネル数: " + channels.length); // チャンネル数をログに出力
  const writeData = [];
  const userList = getUserList();
  for (let i = 0; i < channels.length; i++) {
    // チャンネルの数だけ繰り返す
    const channel = channels[i];
    Logger.log("処理中: #" + channel.name + "(" + channel.id + ")"); // チャンネル名をログに出力
    Logger.log(
      "チャンネルの最終投稿日を取得: #" + channel.name + "(" + channel.id + ")"
    ); // チャンネル名をログに出力
    let historyApi = {};
    historyApi = callSlackApi(
      // チャンネルの履歴を取得
      "conversations.history?channel=" + channel.id
    );
    if (!historyApi.ok) {
      // レスポンスがエラーの場合
      Logger.log("チャンネルに参加: #" + channel.name + "(" + channel.id + ")"); // チャンネル名をログに出力
      callSlackApi("conversations.join?channel=" + channel.id); // チャンネルに参加する
      historyApi = callSlackApi("conversations.history?channel=" + channel.id); // 再度履歴を取得
    }
    const messages = historyApi.messages; // メッセージを取り出す
    const latestActivity = tsConvert(
      // 最新のアクティビティを取得
      messages.filter(
        (message) =>
          !message.subtype ||
          (message.subtype !== "channel_join" &&
            message.subtype !== "channel_leave")
      ).length > 0 // メッセージがある場合
        ? messages.filter(
            (message) =>
              !message.subtype ||
              (message.subtype !== "channel_join" &&
                message.subtype !== "channel_leave")
          )[0].ts
        : channel.created // メッセージがない場合はチャンネル作成日を取得
    );
    const created = tsConvert(channel.created); // チャンネル作成日を取得
    const creator = {
      // チャンネル作成者の情報を取得
      user_id: channel.creator,
      display_name: userList.find((user) => user.user_id === channel.creator)
        .display_name,
      email: userList.find((user) => user.user_id === channel.creator).email,
    };
    writeData.push([
      // チャンネルの情報を配列に格納
      "FALSE",
      channel.id,
      channel.name,
      channel.num_members,
      channel.is_shared,
      creator.display_name,
      creator.email,
      creator.user_id,
      created,
      null,
      latestActivity,
    ]);
    Logger.log("完了: #" + channel.name + "(" + channel.id + ")"); // チャンネル名をログに出力
    Logger.log("待機中..."); // 待機中をログに出力
    Utilities.sleep(3000); // レートリミット対策
  }
  const lastRow = channelSheet.getLastRow();
  if (lastRow > 3) {
    Logger.log("古いデータを削除: " + (lastRow - 3) + "行"); // 古いデータを削除する行数をログに出力
    channelSheet.deleteRows(4, channelSheet.getLastRow() - 3); // チャンネル一覧シートの古いデータを削除
  }
  if (writeData.length > 0) {
    Logger.log("新しいデータを書き込み: " + writeData.length + "行"); // 新しいデータを書き込む行数をログに出力
    channelSheet.insertRowsAfter(3, writeData.length); // 必要に応じて行を作る
    channelSheet
      .getRange(4, 1, writeData.length, writeData[0].length)
      .setValues(writeData); // チャンネル一覧シートに新しいデータを書き込む
  }
  Logger.log("完了"); // ログに出力
}

function archiveButton() {
  const targetChannels = targetSheet // 対象チャンネルを取得
    .getRange("A3:A" + targetSheet.getLastRow())
    .getValues()
    .flat();
  const targetChannelNames = targetSheet
    .getRange("B3:B" + targetSheet.getLastRow())
    .getValues()
    .flat();
  const execId = Utilities.getUuid();
  Logger.log("実行ID: " + execId); // 実行IDをログに出力
  for (let i = 0; i < targetChannels.length; i++) {
    const targetChannel = targetChannels[i];
    const channelName = targetChannelNames[i];
    Logger.log("処理中: " + targetChannel); // 対象チャンネルをログに出力
    const response = archiveChannel(targetChannel);
    if (!response.ok) {
      Logger.log("エラー: (" + targetChannel + ") " + JSON.stringify(response)); // エラーをログに出力
    } else {
      Logger.log(
        "アーカイブ完了: (" + targetChannel + ") " + JSON.stringify(response)
      ); // ログに出力
    }
    writeLog(
      execId,
      response.ok,
      targetChannel,
      channelName,
      JSON.stringify(response)
    ); // スプレッドシートに書き込む
    Logger.log("待機中..."); // 待機中をログに出力
    Utilities.sleep(3000); // レートリミット対策
  }
  // 実行IDをセット
  sheetsApi.getSheetByName("実行結果").getRange("B1").setValue(execId);
  Logger.log("完了"); // ログに出力
}

function archiveChannel(channelId) {
  const archiveApi = callSlackApi("conversations.archive?channel=" + channelId);
  return archiveApi;
}

function callSlackApi(endpoint, payload) {
  Logger.log("Slack APIを呼び出し: " + endpoint); // ログに出力
  const slackApi = UrlFetchApp.fetch("https://slack.com/api/" + endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Authorization: `Bearer ${slackToken}`,
    },
    payload: payload,
  });
  return JSON.parse(slackApi.getContentText());
}

function getUserList() {
  const slackApi = callSlackApi("users.list");
  if (!slackApi.ok) {
    return;
  }
  const userList = slackApi.members
    .filter((user) => !user.is_bot && !user.deleted)
    .map((user) => {
      return {
        user_id: user.id,
        display_name: user.profile.display_name || user.real_name,
        email: user.profile.email,
      };
    });
  return userList;
}

function tsConvert(ts) {
  return Utilities.formatDate(
    new Date(ts * 1000),
    "Asia/Tokyo",
    "yyyy/MM/dd HH:mm:ss"
  ).split(".")[0];
}
