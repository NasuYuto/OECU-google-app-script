//カレンダーから参加する部員の名前・天気を出力する関数
function sendToLine() {
  // LINE Notifyのアクセストークンを設定
  //var accessToken2 = 'LineNotifyのアクセストークン';//幹部
  //var accessToken = 'LineNotifyのアクセストークン';//個人
  var accessToken = 'LineNotifyのアクセストークン';//全体

  var d =  new Date(); // 現在の日付と時刻
  var currentDate = d.getDate()//現在の日付部分を所得
  Logger.log(currentDate)
  var mon = d.getMonth()
  if((mon+=1)==13){//月数字ずれの修正
    mon=1;
  }
  var year = d.getFullYear();  
  var time,night,morning,san,tenkiflag=[0,0,0];
  var cellValues = [];// データを保持するための配列を初期化。この変数は最終的にメッセージとして送るための変数
  var flag = 0 //カレンダーのセルに文字が入っているかどうかの判断に使われる。例えば指定したセルの中に文字が入っていなければそれは通知する必要がないのでflag=1をセットする。
  var range = sheet.getRange(1, 1, 32, 15); // 32行 x 15列 セルの範囲を指定（A1からO32までのセル）
  tenkiflag[0]=0;
  tenkiflag[1]=0;
  tenkiflag[2]=0;  
  
  
  var sheet = SpreadsheetApp.openById("任意のSpredId").getSheetByName(シート名); // シート名を適切に変更
  var sity = '調べたい住所の郵便番号'
  var api= 'OpenWeatherMapのapiキー'
  var requestUrl = 'http://api.openweathermap.org/data/2.5/forecast?zip='+sity+'&units=metric&lang=ja&appid='+api;

  var response = UrlFetchApp.fetch(requestUrl).getContentText();
  var json = JSON.parse(response);
 
  var weatherInfo = [];
  var MIN=json['list'][0]['main']['temp_min'];//最低気温
  var MAX=json['list'][0]['main']['temp_max'];//最高気温
  for(let i=0;i<8;i++){
    weatherInfo[i] = [];
    //Open Weather Mapから取得した天気予報の中から必要な情報を2次元配列に書き込み
    weatherInfo[i][0] = json['list'][i]['dt_txt'];
    time = weatherInfo[i][0].split(" ")
 
   if((time[1]=='06:00:00'||time[1]=='09:00:00') && tenkiflag[0]!=1){
        morning=json['list'][i]['weather'][0]['main'];
        if(morning=='Rain'){//雨が少しでも含まれている場合天気マークを雨にしたいから
          tenkiflag[0]=1;
        }
   }
   else if((time[1]=='012:00:00'||time[1]=='15:00:00')&&tenkiflag[1]!=1){
      san=json['list'][i]['weather'][0]['main'];
      if(san=='Rain'){
        tenkiflag[1]=1;
      }
   }
   
   else if(time[1]=='18:00:00'||time[1]=='21:00:00'){
      night=json['list'][i]['weather'][0]['main'];
      if(night=='Rain'){
        tenkiflag[2]=1;
      }
   }
    weatherInfo[i][1] = json['list'][i]['weather'][0]['main'];//スプレッドシートに出力記録用
    weatherInfo[i][2] = json['list'][i]['main']['temp_min'];
    weatherInfo[i][3] = json['list'][i]['main']['temp_max'];
    weatherInfo[i][4] = json['list'][i]['main']['humidity'];
    weatherInfo[i][5] = json['list'][i]['main']['pressure'];

    //最低気温最高気温のソート
    if(MIN>json['list'][i]['main']['temp_min']){
      MIN=json['list'][i]['main']['temp_min'];
    }
    if(MAX<json['list'][i]['main']['temp_max']){
      MAX=json['list'][i]['main']['temp_max'];
    }
  }
  MIN = (Math.round(MIN*10))/10
  MAX = (Math.round(MAX*10))/10

  //天気情報をスプレッドシートに出力
  var spreadsheet = SpreadsheetApp.openById("任意のSpredId");
  var sheet2 = spreadsheet.getSheetByName('天気予報');
  sheet2.getRange(1,1).setValue('町名')
  sheet2.getRange(2, 1, weatherInfo.length, weatherInfo[0].length).setValues(weatherInfo);


//曜日設定
  var date = new Date()
  var dayWeek = date.getDay();//現在の曜日を所得。日曜日なら0が、月曜なら1が、土曜日なら6が代入される。
  var dayWeekstr = ["日","月","火","水","木","金","土"][dayWeek];

  // セルの値を取得し、日付と比較
  var values = range.getValues();//シートの指定した範囲を2次元配列でvaluesに格納
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
    //rowはセルのy軸colはセルのx軸と考えればよい

      // 日付が条件を満たす場合に通知を送信
      if (currentDate === values[row][col]) { //今日の日付とスプレッドシートに記入された数値が一致した場合次の処理
        if(dayWeekstr==='月'||dayWeekstr==='火'||dayWeekstr==='水'||dayWeekstr==='木'||dayWeekstr==='金' ){ //平日の時の処理
          if(values[row][col+1].length==0.0){ //セルの中身が何もない場合の処理
          //if文でcolではなくcol+1にするかというと、スプレッドシートでは日付の一つ右のセルに名前が記入されているからその値を読むために+1を加えている。
            flag=1
          }
          else{
            var d= values[row][col+=1].replace(/[\r\n]+/g,"");//改行されている部分を横一列の文字列にする。
            var charsets = d.split(",")//カンマがあるごとに分ける,charsetsは一次元配列。
            //並び替え
            charsets = soot(charsets);//上で分けた単語群を並び替えるためにsoot関数に移動
            for(var i=0;i<charsets.length;i++){
            cellValues.push('\n'+charsets[i]) //cellValuesに並び替えた単語を改行しながら代入
            }
          }
        }
        
        if(dayWeekstr === '日'|| dayWeekstr === '土'){//休日のセルだった場合セルの内容を表示するだけ
          if(values[row][col+1].length==0.0){//何もなければフラグをセット
            flag=1
          }
          cellValues.push(values[row][col+1]);
        }
      }
    }
  }

  if(morning=='Rain')morning='☔';
  if(morning=='Clouds')morning='☁';
  if(morning=='Clear')morning='☀';
  if(san=='Rain')san='☔';
  if(san=='Clouds')san='☁';
  if(san=='Clear')san='☀';
  if(night=='Rain')night='☔';
  if(night=='Clouds')night='☁';
  if(night=='Clear')night='☀';

  if((dayWeekstr==='月'||dayWeekstr==='火'||dayWeekstr==='水'||dayWeekstr==='木'||dayWeekstr==='金' ) && flag===0){//平日の場合のメッセージ処理
    sendLineNotification(accessToken,'\n'+ year+'年'+mon+'月'+currentDate+'日'+'\n本日の参加者:\n' +cellValues + '\n\n最低気温: '+MIN+' ℃\n最高気温: ' +MAX+'℃\n\n'+' 朝 | 昼 | 夜\n－－－－－\n'+morning+'|'+san+'|'+night);
  }
  
  if((dayWeekstr === '日'|| dayWeekstr === '土')&& flag===0){//休日の場合のメッセージ処理
    sendLineNotification(accessToken,'\nおはようございます!\n本日は\n'+cellValues+'\n\nです!'+ '\n\n最低気温: '+MIN+' ℃\n最高気温: ' +MAX+'℃\n\n'+' 朝 | 昼 | 夜\n－－－－－\n'+morning+'|'+san+'|'+night);
  }
  
  //sendLineNotification(accessToken,"メンテナンス中")
  Logger.log(cellValues);
}

//幹部ラインに次の日の参加者と二回生以上の人数を表示する関数
function sendToLine2(){
  // LINE Notifyのアクセストークンを設定
  var accessToken2 = 'LineNotifyのアクセストークン';//幹部
  //var accessToken2 = 'LineNotifyのアクセストークン';//個人
  //var accessToken = 'LineNotifyのアクセストークン';//全体

  var d =  new Date(); // 現在の日付と時刻
  var currentDate = d.getDate() //現在の日付部分を所得
  var mon = d.getMonth()
  Logger.log(mon+1)
  
  var sheet = SpreadsheetApp.openById("任意のSpredId").getSheetByName("任意のSpredSheetId"); 
  var d =  new Date(); // 現在の日付と時刻
  var currentDate = d.getDate()
  var cellValues = []; // データを保持するための配列を初期化
  var flag = 0
  // セルの範囲を指定（A1からO32までのセル）
  var range = sheet.getRange(1, 1, 32, 15); // 32行 x 15列
  var date = new Date()
  var dayWeek = date.getDay();
  var dayWeekstr = ["日","月","火","水","木","金","土"][dayWeek];
  Logger.log(dayWeekstr);

  // セルの値を取得し、日付と比較
  var values = range.getValues();
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      //ここまではsendToLine関数と同じ処理
      // 日付が条件を満たす場合に通知を送信
      if (currentDate === values[row][col]) {
        if(dayWeekstr==='日'||dayWeekstr==='月'||dayWeekstr==='火'||dayWeekstr==='水'||dayWeekstr==='木'){//次の日が全員参加じゃない曜日を指定
          if(values[row][col+3].length==0.0){//col+3にすることで次の日のセルを指定できる
            flag=1
          }
          var d= values[row][col+=3].replace(/[\r\n]+/g,"");
          var charsets = d.split(",")
          //並び替え
          charsets = soot(charsets,values[row][col+3]);
          for(var i=0;i<charsets.length;i++){
            cellValues.push('\n'+charsets[i])
          }
          var jou = soot2(charsets)//soot2関数で2回生以上が何人いるかをカウントしjouに格納する
        }
      }
    }
  }
  if((dayWeekstr==='月'||dayWeekstr==='水'||dayWeekstr==='木') && flag===0){
    sendLineNotification(accessToken2, '\n明日の参加者:\n ' + cellValues + '\n\n' + '2年生以上は'+jou+'人です');
  }

  Logger.log(cellValues);

}

//出力する前に学年順にソートを行う関数
function soot(charsets){
 var sheet = SpreadsheetApp.openById("任意のSpredId").getSheetByName("任意のSpredSheetId");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, 2); 
  var values = range.getValues();
  var rchar=[];
  var count =0
   //values配列とcharsets配列を一つずつ見比べて、一致すれば一致した値をrchar配列に格納する。つまりvalues配列の上から見比べているので自動的に学年順に格納される。
  for(var i=0;i<lastRow;i++){
    for(var j=0;j<charsets.length;j++){
      if(values[i][0]==charsets[j]){
        rchar[count] = charsets[j];
        count++;
      }
    }
  }
  Logger.log(values)
  return rchar;
}

//日曜日の時は予定表入力を促すメッセージを送信する関数
function sanday(){
 // LINE Notifyのアクセストークンを設定
  //var accessToken2 = 'LineNotifyのアクセストークン';//幹部
  //var accessToken2 = 'LineNotifyのアクセストークン';//個人
  var accessToken2 = 'LineNotifyのアクセストークン';//全体

  // スプレッドシートのシートとセルの情報を設定
  var d =  new Date(); // 現在の日付と時刻
  var currentDate = d.getDate()
  // データを保持するための配列を初期化
  var cellValues = [];
  var date = new Date()
  var dayWeek = date.getDay();
  var dayWeekstr = ["日","月","火","水","木","金","土"][dayWeek];
  Logger.log(dayWeekstr);
  // セルの値を取得し、日付と比較
 
  if(dayWeekstr==='日' ){
    sendLineNotification(accessToken2, '\n【通知】\n2週間後までの予定表入力をお願いします!');
  }
}

//2回生以上が何人いるかをカウントする関数
function soot2(charsets){
  var sheet = SpreadsheetApp.openById("1srKDoDsifZavX5e4fcVpAqlVo-uS-k_4qGvdHfbuaLE").getSheetByName("配列用部員名簿");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, 2); 
  var values = range.getValues();
  var rchar=[];
   var count =0
  for(var i=0;i<lastRow;i++){
    for(var j=0;j<charsets.length;j++){
      if(values[i][0]==charsets[j] && values[i][1]>=2.0){//valuesの名前とscharsetsが一致してかつ、values配列の学年が2年以上だったらカウント変数を加算していく
        count++;
      }
    }
  }
  return count;
  Logger.log(rchar);
  Logger.log(charsets[2])
}

//カレンダーシート自動追加関数
function addsite(){
  var spreadsheet = SpreadsheetApp.openById("任意のSpredId");
  var sheet2 =spreadsheet.getSheetByName('任意のSpredSheetId');
  var newSheet2 = sheet2.copyTo(spreadsheet);
  var d =  new Date(); // 現在の日付と時刻
  var currentDate = d.getDate() //現在の日付部分を所得
  var monthcount=0;
  var mon = d.getMonth();
  var year = d.getFullYear(); 
  if(mon==11){
    mon=1;
  }
  else{
    mon+=2;
  }
  var sh2= newSheet2.setName(mon+'月予定表');
 
  if(mon== 4 || mon==6 || mon==9 || mon==11){
    monthcount=30;
  }
  else if(mon==1 || mon==3 || mon==5 || mon==7||mon==8 || mon==10 || mon==12 ){
    monthcount=31;
  }
  else if( mon ==2 && year%4==0){
    monthcount=29
  }else {
    monthcount=28;
  }
  
  var month = d.getMonth();  // 1月が0、12月が11
  var day = 1;
  if(month==11){
    year+=1;
    month=1
  }
  else{
    month+=1;
  }

  var dayOfWeek = getDayOfWeek(year, month, day);
  var range = sh2.getRange(2, 2, 1, 15); 
  var values = range.getValues();
  var days = new Array(30);

  for (var i = 0; i < days.length; i++) {
    days[i] = new Array(13);
  }
 
  var count=1;
  var flag=0;
  Logger.log(days)
  var weekcount=0;

  for(i=0;i<=14;i+=2){
    if(values[0][i]==dayOfWeek){
      weekcount=i;
      break;
    }
  }

  for(i=0;i<6;i++){
    for(j=0;j<13;j+=2){
      if(i==0 && flag==0){
        j=weekcount;
        flag=1;
      }
      if(count<=monthcount){
       days[i*5][j]=count;
      }
      else{
        break;
      }
      count+=1;
    }
  }

  if(days[25][0]===null){
    var sheet = spreadsheet.getSheetByName('任意のSpredSheetId');
    var newSheet = sheet.copyTo(spreadsheet);
    sheet2.deleteSheet(sh2);
    var sh = newSheet.setName('任意のSpredSheetId');
    sheet.copyTo(sh);
  }

  console.log("daysの値"+days)
  console.log("元シートの値"+sh2.getRange(3,2,30,13).getValue)
  // var existingData = sh2.getRange(3,2,30,13).getValue();
  // var updatedData = existingData.concat(days);
  // sh2.getRange(3,2,30,13).setValues(updatedData)
  sh2.getRange(3,2,30,13).setValues(days);//dayの中身は日付の数字だけだが、予定欄のところもベースシートの内容を置いてあげれば新たに作成されたカレンダーに予定の内容が書き込まれそう　
}

//曜日を所得
function getDayOfWeek(year, month, day) {
    var daysOfWeek = ['日', '月', '火', '水', '木', '金', '土']
    var date = new Date(year, month , day);
    Logger.log(date.getDay())
    var dayOfWeek = daysOfWeek[date.getDay()];
    return dayOfWeek;
}

function sendToLine2log(){
  console.log("null")
}

function sandaylog(){
  console.log("null")
}

//LINEと連携した関数ここは基本的にいじる必要はない
function sendLineNotification(accessToken, message) {
  var url = 'https://notify-api.line.me/api/notify';
  var headers = {
    'Authorization': 'Bearer ' + accessToken,
  };
  var payload = {
    'message': message,
  };

  var options = {
    'method': 'post',
    'headers': headers,
    'payload': payload,
  };

  // LINEに通知を送信
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}
