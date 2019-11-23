//////  GLOBAL VARIABLES  //////
  // sheet names
var s_setting = '設定';
var s_address = 'フタッフ一覧';
var s_input   = 'シフト入力';
var s_cal     = 'シフト計算';
var s_print   = '印刷';
var s_hday    = '祝日一覧';
var s_SetInfo = 'FormInfo(変更禁止)';
var s_first   = 'NEW';

var s_shop_name       = [5,7];
var s_shop_boss       = [6,7];
var s_shop_detail     = [5,3];
var s_worktime        = [7,3];
var s_opentime        = [8,3];
var s_finishtime      = [9,3];
var s_budget          = [11,3];
var s_workday         = [12,3];
var s_holiday         = [13,3];
var s_resttime        = [9,6];
var s_zero            = [11,6];
var s_half            = [12,6];
var s_one             = [13,6];
var s_onehalf         = [14,6];
var s_divide          = [9,7];
var s_up              = [10,7];
var s_down            = [10,8];
var s_sendurl         = [15,3];

var s_phday           = [17,1];

var s_save_name       = [2,2];
var s_timesheet_year  = [5,4];
var s_timesheet_month = [6,4];
var s_mem_num         = [8,3];
var s_Tsum_index      = [11,3];
var s_btime_index     = [12,3];
var s_mem_index       = [13,3];
var s_Tdiff_index     = [14,3];
var s_id              = [17,3];
var s_shurl           = [18,3];
var s_formsheet       = [19,3];
var s_formurl         = [20,3];
var s_formId          = [21,3];  
var s_folderId        = [22,3];

  // color
var font_col   = 'white';
var setting_ti = '#0080a2';
var setting_in = '#393e4f';
var info_ti    = '#b7282e';
var info_in    = '#95949a';
var hcolor     = '#FF6565';
var hcolor_n   = '#ff6565';
var busy       = '#800000';
var input_bk1  = '#BCBABE';
var input_bk2  = '#87ceeb';

////////////////////////////////////////////////////////////
//////////////////     First Setting      //////////////////
////////////////////////////////////////////////////////////


//////  Add Menu  //////
function onOpen(){
  var sheet   = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [       
    {name:"シフトの作成",functionName:"MakeTimeSheet"},
    {name:"シフト作成メール送信",functionName:"NotifyStaffCreateTS"},
    {name:"シフトの計算",functionName:"CalTimeSheet"},
    {name:"更新",functionName:"UpdateTimeSheet"},
    null,
    {name:"印刷用出力",functionName:"ExportSheetPrint"},
    {name:"PDF化",functionName:"CreatePDF"},
    {name:"完成シフト送信",functionName:"NotifyStaffCompleteTS"},
    null,
    {name:"設定",functionName:"ForAddMenu"},
  ];
    sheet.addMenu("シフト作成", entries);
    }
   
// Add Menu setting
function ForAddMenu(){
  var text = "実行したい操作の数字を半角で入力してください。\\n"
  + "1: 初期設定\\n"
  + "2: アドレス登録画面作成\\n"
  + "3: 登録の削除\\n"
  + "4: ファイルのコピー作成\\n"
  + "5: ファイルの譲渡\\n"
  + "6: 祝日の更新";
  
  var res = Browser.inputBox(text);
  
  if(res === '1'){
    Make_Sheets();
  }
  else if(res === '2'){
    Create_Form();
  }
  else if(res === '3'){
    DeleteData();
  }
  else if(res === '4'){
    Copy();
  }
  else if(res === '5'){
    ChangeOwner();
  }
  else if(res === '6'){
    UpdatePublicHoliday();
  }
  else{
    Browser.msgBox("実行できませんでした。\\n適切な値を入力してください。")
  }
}
    
    
//////  Make Sheets  //////
function Make_Sheets(){
  var res = Browser.msgBox("このスプレッドシート内の既存のデータを全て削除して新しいシフトを作成してもよろしいですか?", Browser.Buttons.OK_CANCEL);
  if(res === "ok"){
    
    // Initialize Spreadsheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    try{ 
      var SS         = sheet.getSheetByName(s_SetInfo);
      var formshName = SS.getRange(s_formsheet[0],s_formsheet[1]+1).getValue();
      var formid     = SS.getRange(s_formId[0],s_formId[1]+1).getValue();
      
      DeleteFormSheet(formshName,formid);
      
    }catch(e){
      ;
    }finally{
      var sheet_url = sheet.getUrl();
      
      var cnt      = sheet.getNumSheets();
      var lasSh    = sheet.getSheets()[cnt-1];
      var lasSName = lasSh.getName();
      var lasS     = sheet.getSheetByName(lasSName);
      SpreadsheetApp.setActiveSheet(lasS);
      sheet.getActiveSheet().setName(s_first);
      
      var ss_new = sheet.getSheetByName(s_first);
      ss_new.clear();
      for(var i=cnt-1;i>=1;i--){
        var sh = sheet.getSheets()[i-1];
        sheet.deleteSheet(sh);
      }
      
      sheet.getActiveSheet().setName(s_setting);
      sheet.insertSheet(s_cal);
      sheet.insertSheet(s_print);
      sheet.insertSheet(s_address);
      sheet.insertSheet(s_hday);
      sheet.insertSheet(s_SetInfo);
      
      var ss_set  = sheet.getSheetByName(s_setting);
      var ss_ad   = sheet.getSheetByName(s_address);
      var ss_adId = ss_ad.getSheetId();
      var ss_cal  = sheet.getSheetByName(s_cal);
      ss_cal.insertColumnsAfter(26,6);
      var ss_inf  = sheet.getSheetByName(s_SetInfo);
      
      GetPublicHoliday(sheet_url);
      ss_inf.getRange(s_shurl[0],s_shurl[1]+1).setValue(sheet_url);
      
      var fileId = sheet.getId();
      var parentFolder = DriveApp.getFileById(fileId).getParents();
      var folderId = parentFolder.next().getId();
      ss_inf.getRange(s_folderId[0],s_folderId[1]+1).setValue(folderId);
      
      // input item
      const arr_other = ['店名','店長','初期設定','メール登録画面URL'];
      const arr_shop  = ['店舗情報','営業時間','開店','閉店'];
      const arr_btime = ['予算時間','平日','休日'];
      const arr_ts    = ['年度','月'];
      const arr_rtime = ['休憩時間','なし','30分','1時間','1時間半','時間区切り','以上','未満'];
      const arr_save  = ['データの保存(変更禁止)','人数','Index','合計時間','予算時間','人数','差異','その他','シフト入力シートID','シートURL','FormSheetName','FormId','FormURL','FolderID'];
      
      // input
      ss_set.getRange(2,2,2,2).merge();
      ss_set.getRange(2,4,2,6).merge();
      ss_set.getRange(s_shop_name[0],s_shop_name[1]).setValue(arr_other[0]);
      ss_set.getRange(s_shop_boss[0],s_shop_name[1]).setValue(arr_other[1]);
      ss_set.getRange(2,2).setValue(arr_other[2]);
      ss_set.getRange(2,2).setFontSize(18);
      ss_set.getRange(s_sendurl[0],s_sendurl[1]).setValue(arr_other[3]);
      
      ss_set.getRange(s_shop_detail[0],s_shop_detail[1],1,2).merge();
      ss_set.getRange(s_shop_detail[0],s_shop_detail[1]).setValue(arr_shop[0]);
      
      ss_set.getRange(s_worktime[0],s_worktime[1],1,2).merge();
      ss_set.getRange(s_worktime[0],s_worktime[1]).setValue(arr_shop[1]);
      ss_set.getRange(s_opentime[0],s_opentime[1]).setValue(arr_shop[2]);
      ss_set.getRange(s_finishtime[0],s_finishtime[1]).setValue(arr_shop[3]);
      
      ss_set.getRange(s_budget[0],s_budget[1],1,2).merge();
      ss_set.getRange(s_budget[0],s_budget[1]).setValue(arr_btime[0]);
      ss_set.getRange(s_workday[0],s_workday[1]).setValue(arr_btime[1]);
      ss_set.getRange(s_holiday[0],s_holiday[1]).setValue(arr_btime[2]);
      
      ss_set.getRange(s_resttime[0],s_resttime[1],2,1).merge();
      ss_set.getRange(s_divide[0],s_divide[1],1,2).merge();
      ss_set.getRange(s_resttime[0],s_resttime[1]).setValue(arr_rtime[0]);
      ss_set.getRange(s_zero[0],s_zero[1]).setValue(arr_rtime[1]);
      ss_set.getRange(s_half[0],s_half[1]).setValue(arr_rtime[2]);
      ss_set.getRange(s_one[0],s_one[1]).setValue(arr_rtime[3]);
      ss_set.getRange(s_onehalf[0],s_onehalf[1]).setValue(arr_rtime[4]);
      ss_set.getRange(s_divide[0],s_divide[1]).setValue(arr_rtime[5]);
      ss_set.getRange(s_up[0],s_up[1]).setValue(arr_rtime[6]);
      ss_set.getRange(s_down[0],s_down[1]).setValue(arr_rtime[7]);
      
      ss_inf.getRange(s_save_name[0],s_save_name[1],2,5).merge();
      ss_inf.getRange(s_save_name[0],s_save_name[1]).setValue(arr_save[0]);
      ss_inf.getRange(s_save_name[0],s_save_name[1]).setFontSize(14);
      ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]).setValue(arr_ts[0]);
      ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]).setValue(arr_ts[1]);
      ss_inf.getRange(s_mem_num[0],s_mem_num[1]).setValue(arr_save[1]);
      ss_inf.getRange(s_Tsum_index[0]-1,s_Tsum_index[1]).setValue(arr_save[2]);
      ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1]).setValue(arr_save[3]);
      ss_inf.getRange(s_btime_index[0],s_btime_index[1]).setValue(arr_save[4]);
      ss_inf.getRange(s_mem_index[0],s_mem_index[1]).setValue(arr_save[5]);
      ss_inf.getRange(s_Tdiff_index[0],s_Tdiff_index[1]).setValue(arr_save[6]);
      ss_inf.getRange(s_id[0]-1,s_id[1]).setValue(arr_save[7]);
      ss_inf.getRange(s_id[0],s_id[1]).setValue(arr_save[8]);
      ss_inf.getRange(s_shurl[0],s_shurl[1]).setValue(arr_save[9]);
      ss_inf.getRange(s_formsheet[0],s_formsheet[1]).setValue(arr_save[10]);
      ss_inf.getRange(s_formurl[0],s_formurl[1]).setValue(arr_save[11]);
      ss_inf.getRange(s_formId[0],s_formId[1]).setValue(arr_save[12]);
      ss_inf.getRange(s_folderId[0],s_folderId[1]).setValue(arr_save[13]);
      
      ss_inf.getRange(s_mem_num[0],s_mem_num[1]+1,1,2).merge();
      ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1]+1,1,2).merge();
      ss_inf.getRange(s_btime_index[0],s_btime_index[1]+1,1,2).merge();
      ss_inf.getRange(s_mem_index[0],s_mem_index[1]+1,1,2).merge();
      ss_inf.getRange(s_Tdiff_index[0],s_Tdiff_index[1]+1,1,2).merge();
      ss_inf.getRange(s_id[0],s_id[1]+1,1,2).merge();
      ss_inf.getRange(s_shurl[0],s_shurl[1]+1,1,2).merge();
      ss_inf.getRange(s_formsheet[0],s_formsheet[1]+1,1,2).merge();
      ss_inf.getRange(s_formurl[0],s_formurl[1]+1,1,2).merge();
      ss_inf.getRange(s_formId[0],s_formId[1]+1,1,2).merge();
      ss_inf.getRange(s_folderId[0],s_folderId[1]+1,1,2).merge();
      
      // Address sheet
      var ss_ad = sheet.getSheetByName(s_address);
      ss_ad.getRange(1,1).setValue('名前');
      ss_ad.getRange(1,2).setValue('メールアドレス');
      
      // Layout
      // arrangement
      ss_set.getRange(1,1,20,10).setHorizontalAlignment('center');
      ss_set.getRange(1,1,20,10).setVerticalAlignment('middle');
      ss_set.setColumnWidth(1,21);
      ss_set.setColumnWidth(2,21);
      ss_set.setColumnWidth(5,21);
      ss_set.setColumnWidth(6,100);
      ss_set.setColumnWidth(9,21);
      
      ss_inf.getRange(1,1,30,10).setHorizontalAlignment('center');
      ss_inf.getRange(1,1,30,10).setVerticalAlignment('middle');
      ss_inf.setColumnWidth(1,21);
      ss_inf.setColumnWidth(2,14);
      ss_inf.setColumnWidth(3,125);
      ss_inf.setColumnWidth(6,14);
      
      // set
      // setting
      ss_set.getRange(2,2,15,8).setBorder(true,true,true,true,false,false);
      ss_set.getRange(s_shop_detail[0],s_shop_detail[1]).setBorder(true,true,true,true,true,true);
      ss_set.getRange(s_shop_name[0],s_shop_name[1],2,2).setBorder(true,true,true,true,true,true);
      ss_set.getRange(s_resttime[0],s_resttime[1],6,3).setBorder(true,true,true,true,true,true);
      ss_set.getRange(s_worktime[0],s_worktime[1],3,2).setBorder(true,true,true,true,true,true);
      ss_set.getRange(s_budget[0],s_budget[1],3,2).setBorder(true,true,true,true,true,true);
      ss_set.getRange(s_sendurl[0],s_sendurl[1],1,2).setBorder(true,true,true,true,true,true);
      // SetInfo
      ss_inf.getRange(s_save_name[0],s_save_name[1],22,5).setBorder(true,true,true,true,false,false);
      ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]-1,2,2).setBorder(true,true,true,true,true,true);
      ss_inf.getRange(s_mem_num[0],s_mem_num[1],1,2).setBorder(true,true,true,true,true,true);
      ss_inf.getRange(s_Tsum_index[0]-1,s_Tsum_index[1]).setBorder(true,true,true,true,true,true);
      ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1],4,3).setBorder(true,true,true,true,true,true);
      ss_inf.getRange(s_id[0]-1,s_id[1]).setBorder(true,true,true,true,true,true);
      ss_inf.getRange(s_id[0],s_id[1],6,3).setBorder(true,true,true,true,true,true);
      
      // color
      // setting
      ss_set.getRange(2,2,2,8).setBackground(setting_ti);
      ss_set.getRange(s_shop_name[0],s_shop_name[1],2,1).setBackground(setting_in);
      ss_set.getRange(s_shop_detail[0],s_shop_detail[1]).setBackground(setting_in);
      ss_set.getRange(s_worktime[0],s_worktime[1]).setBackground(setting_in);
      ss_set.getRange(s_opentime[0],s_opentime[1],2,1).setBackground(setting_in);
      ss_set.getRange(s_budget[0],s_budget[1]).setBackground(setting_in);
      ss_set.getRange(s_workday[0],s_workday[1],2,1).setBackground(setting_in);
      ss_set.getRange(s_resttime[0],s_resttime[1],2,3).setBackground(setting_in);
      ss_set.getRange(s_zero[0],s_zero[1],4,1).setBackground(setting_in);
      ss_set.getRange(s_sendurl[0],s_sendurl[1]).setBackground('#2f5d50');
      
      ss_set.getRange(2,2).setFontColor(font_col);
      ss_set.getRange(s_shop_detail[0],s_shop_detail[1]).setFontColor(font_col);
      ss_set.getRange(s_shop_name[0],s_shop_name[1],2,1).setFontColor(font_col);
      ss_set.getRange(s_worktime[0],s_worktime[1]).setFontColor(font_col);
      ss_set.getRange(s_opentime[0],s_opentime[1],2,1).setFontColor(font_col);
      ss_set.getRange(s_budget[0],s_budget[1]).setFontColor(font_col);
      ss_set.getRange(s_workday[0],s_workday[1],2,1).setFontColor(font_col);
      ss_set.getRange(s_resttime[0],s_resttime[1],2,3).setFontColor(font_col);
      ss_set.getRange(s_zero[0],s_zero[1],4,1).setFontColor(font_col);
      ss_set.getRange(s_sendurl[0],s_sendurl[1]).setFontColor(font_col);
      
      // address
      ss_ad.setColumnWidth(1,115);
      ss_ad.setColumnWidth(2,324);
      
      // SetInfo
      ss_inf.getRange(s_save_name[0],s_save_name[1]).setBackground(info_ti);
      ss_inf.getRange(s_Tsum_index[0]-1,s_Tsum_index[1]).setBackground(info_ti);
      ss_inf.getRange(s_id[0]-1,s_id[1]).setBackground(info_ti);
      ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1],2,1).setBackground(info_in);
      ss_inf.getRange(s_mem_num[0],s_mem_num[1]).setBackground(info_in);
      ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1],4,1).setBackground(info_in);
      ss_inf.getRange(s_id[0],s_id[1],6,1).setBackground(info_in);
      
      ss_set.setHiddenGridlines(true);
      ss_inf.setHiddenGridlines(true);
      
    }
  }
}
  



////////////////////////////////////////////////////////////
//////////////////     Google Form        //////////////////
////////////////////////////////////////////////////////////  


//////  Create New Google Form  //////
function Create_Form(){
  var res = Browser.msgBox("新しくメール登録フォームを作成しますか?", Browser.Buttons.OK_CANCEL);
  if(res === "ok"){
    var sh     = SpreadsheetApp.getActiveSpreadsheet();
    var ss_set = sh.getSheetByName(s_setting);
    var ss_inf = sh.getSheetByName(s_SetInfo);
    var ssid   = sh.getId();
    var form   = FormApp.create('メール登録');
    form.setTitle('メールアドレスを登録してください');
    form.setDescription('名前、メールアドレスを入力してください．\n' + 'メールアドレスは何でも構いません.');
    form.addTextItem().setTitle('氏名 (例 田中太郎)').setRequired(true);
    form.addTextItem().setTitle('メールアドレス').setRequired(true);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ssid);
    
    var trigger = ScriptApp.newTrigger('ExportAddress').forSpreadsheet(ssid).onFormSubmit().create();
    
    var SendUrl = form.getPublishedUrl();
    var EditUrl = form.getEditUrl();
    var saveId  = form.getId();
    ss_inf.getRange(s_formurl[0],s_formurl[1]+1).setValue(EditUrl);
    ss_inf.getRange(s_formId[0],s_formId[1]+1).setValue(saveId);
    ss_set.getRange(s_sendurl[0],s_sendurl[1]+1).setValue(SendUrl);
    SpreadsheetApp.flush();
    Browser.msgBox('以下のURLをスタッフに送信してください．\\n'+ SendUrl);
    var shName     = sh.getSheets()[0];
    var BindShName = shName.getName();
    ss_inf.getRange(s_formsheet[0],s_formsheet[1]+1).setValue(BindShName);
  }
}


//////  Submit Action  //////
// **トリガーの設定と、反映するスプレッドシートの設定をする必要あり．**
function ExportAddress(e){
  var itemResponses = e.values; // Spreadsheet側でのスクリプトなのでvaluesでok.ハマった．
  
  var sheet  = SpreadsheetApp.getActiveSpreadsheet();
  var ss     = sheet.getSheetByName(s_address);
  var ss_set = sheet.getSheetByName(s_setting);
  
  var l_r      = ss.getLastRow();
  var add_name = itemResponses[1];
  var add_mail = itemResponses[2];
  ss.getRange(l_r+1,1).setValue(add_name);
  ss.getRange(l_r+1,2).setValue(add_mail); 
  
  Reply_Mail(add_name,add_mail);
}


//////  Send Complete Mail  //////
function Reply_Mail(add_name,add_mail){
  var ss      = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s_address);
  var new_row = ss.getLastRow();
  var name    = add_name;
  var mail    = add_mail;
  
  var title = '登録完了';
  var contents 
  = '登録が以下のように完了しました．\n'
  + '---------------------------\n'
  + '名前 : '+ name + '\n\n'
  + 'メールアドレス : ' + mail + '\n'
  + '---------------------------\n';
  
  GmailApp.sendEmail(mail, title, contents);
}


//////  Delete Data  //////
function DeleteData(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss_ad = sheet.getSheetByName(s_address);
  
  var delete_name = Browser.inputBox("削除する人の名前をフルネームで入力してください．");
  var row   = ss_ad.getLastRow();
  var index = 2;
  
  for(var i=0;i<row-1;i++){
    var name  = ss_ad.getRange(2+i,1).getValue();
    var check = name.indexOf(delete_name);
    if(check !== -1){
      del_name = ss_ad.getRange(index,1).getValue();
      ss_ad.deleteRow(index);
      break;
    }
    index = index+1;
  }
  Browser.msgBox(del_name+'さんのデータを正常に削除しました.');   
  
}




////////////////////////////////////////////////////////////
//////////////////      Time Sheet        //////////////////
////////////////////////////////////////////////////////////


//////  Make Timesheet  //////
function MakeTimeSheet() {
  var res = Browser.msgBox("新しいシフト入力シートを作成しますか?", Browser.Buttons.OK_CANCEL);
  if(res === 'ok'){
    var r      = Browser.inputBox("作成する月を入力してください．");
    var sheet  = SpreadsheetApp.getActiveSpreadsheet();
    var ss_set = sheet.getSheetByName(s_setting);
    var ss_ad  = sheet.getSheetByName(s_address);
    var ss_inf = sheet.getSheetByName(s_SetInfo);
    var ss_cal = sheet.getSheetByName(s_cal);
    ss_cal.clear();
    
    ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]-1).setValue(r);
    
    // Input Spreadsheet
    var folderId  = ss_inf.getRange(s_folderId[0],s_folderId[1]+1).getValue();
    var nameShop  = ss_set.getRange(s_shop_name[0],s_shop_name[1]+1).getValue();
    var yearT     = ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]-1).getValue();
    var monthT    = ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]-1).getValue();
    
    // MAKE INPUT SHEET //
    // Get data
    var DATE,Shop_time,Pday,step1,i;
    var ALL   = GetSetting();
    DATE      = ALL[0];
    Shop_time = ALL[1];
    Pday      = ALL[2];
    var day   = DATE[0].getDate();
    var num   = DATE[1].getDate() - (day+1+15+1);
    
    var s_inp = CreateSpreadsheetInfolder(folderId,nameShop+'_'+monthT+'月シフト');
    var s_inputId = s_inp.getId();
    var ss_i      = SpreadsheetApp.openById(s_inputId);
    ss_i.insertColumnsAfter(26,7);
    ss_i.getActiveSheet().setName(s_input);
    var ss_input  = ss_i.getSheetByName(s_input);
    ss_input.setHiddenGridlines(true);
    ss_inf.getRange(s_id[0],s_id[1]+1).setValue(s_inputId);
    
    var mem_num = ss_ad.getRange(2,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() - 1;
    var member  = ss_ad.getRange(2,1,mem_num,1).getValues();
    for(var i=0;i<mem_num;i++){
      step1 = i*2;
      ss_input.getRange(step1+7,1,2,1).merge();
      ss_input.getRange(step1+7,1).setValue(member[i][0]);
    }
    
    // set Day  
    // this month
    var dweek_index1 = DATE[0].getDay(); // 2 曜日のインデックス
    var dweek = ['日', '月', '火', '水', '木', '金', '土'];
    var count = 0;
    var c;
    var index;
    for(i=0;i<=DATE[1].getDate()-day;i++){ // 月の終わりから16日を引く
      index = (dweek_index1+i) % 7;
      ss_input.getRange(5,2+i).setValue(day+i);
      ss_input.getRange(6,2+i).setValue(dweek[index]);
      ss_cal.getRange(5,2+i).setValue(day+i);
      ss_cal.getRange(6,2+i).setValue(dweek[index]);
      if(Pday.indexOf(Utilities.formatDate(new Date(DATE[0].getYear(), DATE[0].getMonth(), day+i), "JST", "yyyy/MM/dd")) !== -1 ||
         (dweek[index] === dweek[0]) || 
        (dweek[index] === dweek[6])){
          ss_input.getRange(5,2+i).setBackground(hcolor);
          ss_input.getRange(6,2+i).setBackground(hcolor);
          ss_cal.getRange(5,2+i).setBackground(hcolor);
          ss_cal.getRange(6,2+i).setBackground(hcolor);
        }
      count = count + 1;
    }
    
    // next month
    c = count
    day = DATE[2].getDate();
    var dweek_index2 = DATE[2].getDay();
    for(i=1;i<=15;i++){
      index = (dweek_index2+i-1) % 7;
      ss_input.getRange(5,c+1+i).setValue(day+(i-1));
      ss_input.getRange(6,c+1+i).setValue(dweek[index]);
      ss_cal.getRange(5,c+1+i).setValue(day+(i-1));
      ss_cal.getRange(6,c+1+i).setValue(dweek[index]);
      if(Pday.indexOf(Utilities.formatDate(new Date(DATE[2].getYear(), DATE[2].getMonth(), day+(i-1)), "JST", "yyyy/MM/dd")) !== -1 ||
        (dweek[index] === dweek[0]) ||
          (dweek[index] === dweek[6])){
            ss_input.getRange(5,c+1+i).setBackground(hcolor);
            ss_input.getRange(6,c+1+i).setBackground(hcolor);
            ss_cal.getRange(5,c+1+i).setBackground(hcolor);
            ss_cal.getRange(6,c+1+i).setBackground(hcolor);
          }
      count = count + 1;
    }
   
    // Layout
    ss_input.getRange(3,1,2,5).merge();
    ss_input.getRange(3,1).setValue(DATE[0].getYear()+'/'+(DATE[0].getMonth()+1)+'/16 ~ '+DATE[2].getYear()+'/'+(DATE[2].getMonth()+1)+'/15'); // getMonthは+1
    ss_input.getRange(3,1).setFontSize('16');
    ss_input.getRange(5,1,2,1).merge();
    
    ss_cal.getRange(3,1,2,5).merge();
    ss_cal.getRange(3,1).setValue(DATE[0].getYear()+'/'+(DATE[0].getMonth()+1)+'/16 ~ '+DATE[2].getYear()+'/'+(DATE[2].getMonth()+1)+'/15'); // getMonthは+1
    ss_cal.getRange(3,1).setFontSize('16');
    ss_cal.getRange(5,1,2,1).merge();
    
    //Background color
    var col = count + 1; // day + name
    for(i=0;i<mem_num;i++){
      step1 = i*2;
      if(i % 2 === 0){
        ss_input.getRange(7+step1,1,2,col).setBackground(input_bk1);
      }else if(i % 2 === 1){
        ss_input.getRange(9+(step1-2),1,2,col).setBackground(input_bk2);
      }
    }
    
    // each cells
    ss_input.getRange(1,1,(mem_num*2)+7,col).setHorizontalAlignment('center');
    ss_input.getRange(1,1,(mem_num*2)+7,col).setVerticalAlignment('middle');
    ss_input.setColumnWidths(1,1,70);
    ss_input.setColumnWidths(2, col-2+1, 50);
    ss_input.getRange(5,1,(mem_num*2)+2,col).setBorder(true,true,true,true,true,true);
    
    // Make Dropdown
    DropDown(ss_input,7,2,mem_num*2,count);
    
    
    // MAKE CALCULATION SHEET //
    // layout
    const c1 = "#B2B2B2";
    const c2 = "#99B2FF";
    for(i=0;i<mem_num;i++){
      step1 = 4*i;
      ss_cal.getRange(7+step1,1).setValue('出勤時間');
      ss_cal.getRange(7+step1,1,1,col).setBackground(c1);
      ss_cal.getRange(8+step1,1).setValue('休憩');
      ss_cal.getRange(8+step1,1,1,col).setBackground(c2);
      ss_cal.getRange(9+step1,1,2,1).merge();
      ss_cal.getRange(9+step1,1).setValue(member[i][0]);
    }
    
    const mem_index   = 10+step1+1;
    const btime_index = 10+step1+2;
    const Tsum_index  = 10+step1+3;
    const Tdiff_index = 10+step1+4;
    const c3 ="#B2B2B2";
    const c4 ="#65CB00";
    const c5 ="#99CCFF";
    ss_cal.getRange(mem_index,1).setValue("人数");
    ss_cal.getRange(mem_index,1,1,col).setBackground(c3);
    ss_cal.getRange(btime_index,1).setValue("予算時間");
    ss_cal.getRange(btime_index,1,1,col).setBackground(c4);
    ss_cal.getRange(Tsum_index,1).setValue("合計時間");
    ss_cal.getRange(Tsum_index,1,1,col).setBackground(c5);
    ss_cal.getRange(Tdiff_index,1).setValue("時間差異");
    ss_cal.getRange(Tdiff_index,1,1,col).setBackground(c4);
    
    ss_cal.setColumnWidths(1,1,70);
    ss_cal.setColumnWidths(2, col-2+1, 50);
    ss_cal.getRange(1,1,step1+8+6,col).setHorizontalAlignment('center');
    ss_cal.getRange(1,1,step1+8+6,col).setVerticalAlignment('middle');
    ss_cal.getRange(5,1,step1+8+2,col).setBorder(true,true,true,true,true,true);
    
    // save data
    ss_inf.getRange(s_mem_num[0],s_mem_num[1]+1).setValue(mem_num);
    ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1]+1).setValue(Tsum_index);
    ss_inf.getRange(s_btime_index[0],s_btime_index[1]+1).setValue(btime_index);
    ss_inf.getRange(s_mem_index[0],s_mem_index[1]+1).setValue(mem_index);
    ss_inf.getRange(s_Tdiff_index[0],s_Tdiff_index[1]+1).setValue(Tdiff_index);
    
  }else{
    Browser.msgBox("中断しました．");
  }
}
    

//////  Calculate Time Sheet  //////
function  CalTimeSheet(){
  var res = Browser.msgBox("シフトの計算を行いますか?", Browser.Buttons.OK_CANCEL);
  if(res === "ok"){
    var sheet     = SpreadsheetApp.getActiveSpreadsheet();
    var ss1       = sheet.getSheetByName(s_setting);
    var ss3       = sheet.getSheetByName(s_cal);
    var ss_inf    = sheet.getSheetByName(s_SetInfo);
    var s_inputId = ss_inf.getRange(s_id[0],s_id[1]+1).getValue();
    var ss_i      = SpreadsheetApp.openById(s_inputId);
    var ss2       = ss_i.getSheetByName(s_input);
    
    var c            = ss3.getLastColumn();
    var ALL          = GetSetting();
    var mem_num      = ss_inf.getRange(s_mem_num[0],s_mem_num[1]+1).getValue();
    var Tsum_index   = ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1]+1).getValue();
    var btime_index  = ss_inf.getRange(s_btime_index[0],s_btime_index[1]+1).getValue();
    var mem_index    = ss_inf.getRange(s_mem_index[0],s_mem_index[1]+1).getValue();
    var Tdiff_index  = ss_inf.getRange(s_Tdiff_index[0],s_Tdiff_index[1]+1).getValue();
    var extract_data = ss2.getRange(7,2,mem_num*2,c-1).getValues();
    var input_data   = [];
    
    for(var i=0;i<extract_data.length;i++){
      var pre = extract_data[i];
      for(var p=0;p<extract_data[0].length;p++){
        if((pre[p] !== "") && (pre[p] !== "×")){
          var h = parseFloat(Utilities.formatDate(pre[p], "JST", "HH"));
          var m = parseFloat(Utilities.formatDate(pre[p], "JST", "mm"));
          pre[p] = h + (m / 60);}
      }
      input_data.push(pre);
    }
    
    var DATE      = ALL[0];
    var Shop_time = ALL[1];
    var Pday      = ALL[2];
    var rest_time = Shop_time[2];
    
    var Go_w,Rv_w,step1,step2,step3,s,intime,rtime,bTime,Tdiff;
    var sum_people  = [];
    var sum_time    = [];
    var rest_budget = [];
    var People      = [];
    var Time        = [];
    var Rest        = [];
    
    // Calculation
    for(i=0;i<c-1;i++){
      var sum_n = 0;
      var sum_t = 0;
      
      if(ss3.getRange(6,2+i).getBackground() === hcolor_n){
        Tdiff = Shop_time[1][1][0];
      }else if(ss3.getRange(6,2+i).getBackground() !== hcolor_n){
        Tdiff = Shop_time[1][0][0];
      }
      ss3.getRange(btime_index,i+2).setValue(Tdiff);
      
      for(var k=0;k<(input_data.length/2);k++){
        step1 = k*2;
        s = k*4;
        if(input_data[step1][i]==="×"){
          ss3.getRange(9+s,2+i,2,1).setBackground(busy);
          continue;
        }else if(input_data[step1][i]===""){
          continue;
        }else{
          Go_w = input_data[step1][i]; // 想定：時間の入力は9or9.5など
          Rv_w = input_data[step1+1][i];
        }
        intime = Rv_w - Go_w;
        
        if((rest_time[2][0]<=intime) && (intime<rest_time[3][0])){
          sum_n = sum_n+1;
          rtime = 1.0;
          sum_t = sum_t + (intime - rtime);
          Tdiff = Tdiff - (intime - rtime);
        }else if((rest_time[1][0]<=intime) && (intime<rest_time[2][0])){
          sum_n = sum_n+1;
          rtime = 0.5;
          sum_t = sum_t + (intime - rtime);
          Tdiff = Tdiff - (intime - rtime);
        }else if((0<intime) && (intime<rest_time[1][0])){
          sum_n = sum_n+1;
          rtime = 0.0;
          sum_t = sum_t + (intime - rtime);
          Tdiff = Tdiff - (intime - rtime);
        }else if(rest_time[3][0]<=intime){
          sum_n = sum_n+1;
          rtime = 1.5;
          sum_t = sum_t + (intime - rtime);
          Tdiff = Tdiff - (intime - rtime); 
        }
        
        // set results
        ss3.getRange(7+s,2+i).setValue(intime);
        ss3.getRange(8+s,2+i).setValue(rtime);
        
      }
      sum_people.push(sum_n); // 各列ごとの配列の完成
      sum_time.push(sum_t);
      rest_budget.push(Tdiff);
      
    }
    
    // Paste results
    for(i=0;i<(input_data.length/2);i++){
      step2 = i*2;
      step3 = i*4;
      var emb = [input_data[step2],input_data[step2+1]];
      ss3.getRange(9+step3,2,2,c-1).setValues(emb);
    }
    ss3.getRange(mem_index,2,1,c-1).setValues(People=[sum_people]);
    ss3.getRange(Tsum_index,2,1,c-1).setValues(Time=[sum_time]);
    ss3.getRange(Tdiff_index,2,1,c-1).setValues(Rest=[rest_budget]);
    
    // Check
    var color_alart = '#FFFF00';
    for(i=0;i<c-1;i++){
      var r_d = ss3.getRange(Tdiff_index,2+i);
      var r_t = ss3.getRange(Tsum_index,2+i);
      var r_b = ss3.getRange(btime_index,2+i);
      var t = r_t.getValue();
      var b = r_b.getValue();
      if(t!==b){
        r_d.setBackground(color_alart);
      }
    }
    
  }
}


//////  For Update  //////
function UpdateTimeSheet(){
  var sheet  = SpreadsheetApp.getActiveSpreadsheet();
  var ss1    = sheet.getSheetByName(s_setting);
  var ss3    = sheet.getSheetByName(s_cal);
  var ss_inf = sheet.getSheetByName(s_SetInfo);
  
  var c           = ss3.getLastColumn();
  var ALL         = GetSetting();
  var mem_num     = ss_inf.getRange(s_mem_num[0],s_mem_num[1]+1).getValue();
  var Tsum_index  = ss_inf.getRange(s_Tsum_index[0],s_Tsum_index[1]+1).getValue();
  var btime_index = ss_inf.getRange(s_btime_index[0],s_btime_index[1]+1).getValue();
  var mem_index   = ss_inf.getRange(s_mem_index[0],s_mem_index[1]+1).getValue();
  var Tdiff_index = ss_inf.getRange(s_Tdiff_index[0],s_Tdiff_index[1]+1).getValue();
 
  var DATE      = ALL[0];
  var Shop_time = ALL[1];
  var Pday      = ALL[2];
  var rest_time = Shop_time[2];
  var Tdiff     = ss3.getRange(btime_index,2,1,c-1).getValues();
  Tdiff         = Tdiff[0];

  var Go_w,Rv_w,step1,s,intime,rtime,bTime;
  var sum_people  = [];
  var sum_time    = [];
  var rest_budget = [];
  var People      = [];
  var Time        = [];
  var Rest        = [];
  
  var update_data = ss3.getRange(7,2,4*mem_num,c-1).getValues();
  
  // Calculation
  for(var i=0;i<c-1;i++){
    var sum_n = 0;
    var sum_t = 0;
    
    for(var k=0;k<mem_num;k++){
      step1 = k*4+2;
      //ss3.getRange(7+step1-2,2,2,c-1).clearContent();
      if(update_data[step1][i]==="×"){
        ss3.getRange(7+step1,2+i,2,1).setBackground(busy);
        continue;
      }else if(update_data[step1][i]===""){
        continue;
      }else{
        Go_w = update_data[step1][i]; // 想定：時間の入力は9or9.5など
        Rv_w = update_data[step1+1][i];
      }
      intime = Rv_w - Go_w;
      
      if((rest_time[2][0]<=intime) && (intime<rest_time[3][0])){
        sum_n = sum_n+1;
        rtime = 1.0;
        sum_t = sum_t + (intime - rtime);
        Tdiff[i] = Tdiff[i] - (intime - rtime);
      }else if((rest_time[1][0]<=intime) && (intime<rest_time[2][0])){
        sum_n = sum_n+1;
        rtime = 0.5;
        sum_t = sum_t + (intime - rtime);
        Tdiff[i] = Tdiff[i] - (intime - rtime);
      }else if((0<intime) && (intime<rest_time[1][0])){
        sum_n = sum_n+1;
        rtime = 0.0;
        sum_t = sum_t + (intime - rtime);
        Tdiff[i] = Tdiff[i] - (intime - rtime);
      }else if(rest_time[3][0]<=intime){
        sum_n = sum_n+1;
        rtime = 1.5;
        sum_t = sum_t + (intime - rtime);
        Tdiff[i] = Tdiff[i] - (intime - rtime); 
      }
      
      // set results
      ss3.getRange(7+step1-2,2+i).setValue(intime);
      ss3.getRange(8+step1-2,2+i).setValue(rtime);
      
    }
    sum_people.push(sum_n); // 各列ごとの配列の完成
    sum_time.push(sum_t);
    rest_budget.push(Tdiff[i]);
    
  }
  
  // Paste results
  ss3.getRange(mem_index,2,1,c-1).setValues(People=[sum_people]);
  ss3.getRange(Tsum_index,2,1,c-1).setValues(Time=[sum_time]);
  ss3.getRange(Tdiff_index,2,1,c-1).setValues(Rest=[rest_budget]);
  
  // Check
  var color_alart = '#FFFF00';
  var color_ok    = '#65CB00';
  for(i=0;i<c-1;i++){
    var r_d = ss3.getRange(Tdiff_index,2+i);
    var r_t = ss3.getRange(Tsum_index,2+i);
    var r_b = ss3.getRange(btime_index,2+i);
    var t   = r_t.getValue();
    var b   = r_b.getValue();
    if(t!==b){
      r_d.setBackground(color_alart);
    }else{
      r_d.setBackground(color_ok);
    }
  }

}




////////////////////////////////////////////////////////////
//////////////////          PDF           //////////////////
////////////////////////////////////////////////////////////


//////  For Printing Sheet  //////
function ExportSheetPrint(){
  var sheet  = SpreadsheetApp.getActiveSpreadsheet();
  var ss1    = sheet.getSheetByName(s_setting);
  var ss3    = sheet.getSheetByName(s_cal);
  var ss4    = sheet.getSheetByName(s_print);
  var ss_ad  = sheet.getSheetByName(s_address);
  var ss_inf = sheet.getSheetByName(s_SetInfo);
  ss4.clear();
  
  var mem_num = ss_inf.getRange(s_mem_num[0],s_mem_num[1]+1).getValue();
  var member  = ss_ad.getRange(2,1,mem_num,1).getValues();
  var c       = ss3.getLastColumn();
  var r       = ss3.getLastRow();
  var data    = ss3.getRange(7,2,r-4-6,c-1).getValues();
  var n_data  = [];
  
  
  for(var i=0;i<mem_num;i++){
    var pro_data1 = [];
    var pro_data2 = [];
    var step = (i*4)+2;
    for(var k=0;k<data[0].length;k++){
      if((data[step][k]!=="") && (data[step][k]!=="×")){
        
        var integer_part1 = String(data[step][k]).split(".")[0];
        var decimal_part1 = parseFloat("0."+String(data[step][k]).split(".")[1]);
        if(decimal_part1 === 0.0){
          pro_data1.push(String(integer_part1)+":00");
        }else{
          var chan = decimal_part1*6*10;
          pro_data1.push(integer_part1+":"+chan);
        }
        
        var integer_part2 = String(data[step+1][k]).split(".")[0];
        var decimal_part2 = parseFloat("0."+String(data[step+1][k]).split(".")[1]);
        if(decimal_part2 === 0.0){
          pro_data2.push(String(integer_part2)+":00");
        }else{
          var chan = decimal_part2*6*10;
          pro_data2.push(integer_part2+":"+chan);
        }
        
      }else{
        pro_data1.push("");
        pro_data2.push("");
      }
    }
    n_data.push(pro_data1);
    n_data.push(pro_data2);
  }
  
  for(var i=0;i<mem_num;i++){
    step = i*2;
    ss4.getRange(step+5,1,2,1).merge();
    ss4.getRange(step+5,1).setValue(member[i][0]);
    if(i%2 === 0){
      ss4.getRange(5+step,1,2,c).setBackground(input_bk1);
    }else{
      ss4.getRange(5+step,1,2,c).setBackground(input_bk2);
    }
  }
  
  ss3.getRange(3,1,4,c).copyTo(ss4.getRange(1,1,4,c));
  ss4.getRange(5,2,mem_num*2,c-1).setValues(n_data);
  
  // Layout
  ss4.setColumnWidths(2, c-1, 50);
  ss4.getRange(5,1,mem_num*2,c).setBorder(true,true,true,true,true,true);
  ss4.getRange(5,1,mem_num*2,c).setHorizontalAlignment('center');
  ss4.getRange(5,1,mem_num*2,c).setVerticalAlignment('middle');
  
}

// Create PDF
function CreatePDF(){
  var res = Browser.msgBox("PDFに出力しますか?", Browser.Buttons.OK_CANCEL);
  if(res === "ok"){
    var sheet  = SpreadsheetApp.getActiveSpreadsheet();
    var ss1    = sheet.getSheetByName(s_setting);
    var ss4    = sheet.getSheetByName(s_print);
    var ss_inf = sheet.getSheetByName(s_SetInfo);
    
    var shop     = ss1.getRange(s_shop_name[0],s_shop_name[1]+1).getValue();
    var month    = ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]-1).getValue();
    var folderId = ss_inf.getRange(s_folderId[0],s_folderId[1]+1).getValue();
    var c        = ss4.getLastColumn();
    var r        = ss4.getLastRow();
    var las      = ss4.getRange(1, c);
    las = las.getA1Notation();
    las = las.replace(/\d/,'');
    var last    = las+r;

    var pdfname = shop+'_'+month+'月_シフト.pdf';
    // Create PDF
    //var root= DriveApp.getRootFolder();
    //var folderid = root.getId();
    var ssid = sheet.getId();
    var sheetid = ss4.getSheetId();
    var timestamp = getTimestamp();
    var p_range = ss4.getRange(ss4.getLastRow(),ss4.getLastColumn());
    
    var folder = DriveApp.getFolderById(folderId);
    
    // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
    var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
    
    // Option of PDF
    var opts = {
      exportFormat: "pdf",
      format:       "pdf",
      size:         "A4",
      top_margin:   "0.25",
      bottom_margin:"0.19",
      left_margin:  "0.20",
      right_margin: "0.20",
      range:        "A1%3A"+last,
      portrait:     "false", // if true, vertical direction.
      fitw:         "true",
      sheetnames:   "false",
      printtitle:   "false",
      pagenumbers:  "false",
      gridlines:    "false",
      fzr:          "false",
      gid:          sheetid};
    
    var url_ext = [];
    
    for( optName in opts ){
      url_ext.push( optName + "=" + opts[optName] );
    }
    
    var options = url_ext.join("&");
    
    // API使用のためのOAuth認証
    var token = ScriptApp.getOAuthToken();
    
    var response = UrlFetchApp.fetch(url + options, {headers: {'Authorization': 'Bearer ' +  token}});
    var blob = response.getBlob().setName(pdfname);    
    folder.createFile(blob);
  }
}




//////////////////////////////
//////   Notify Staff   //////
//////////////////////////////

// Send Create Time Sheet
function NotifyStaffCreateTS(){
  var res = Browser.inputBox('共有するスプレッドシートのリンクを貼り付けてください。');
  var lim = Browser.inputBox('期限は何日までにしますか?');
  if((res !== "cansel") || (lim !== "cansel")){
    var sheet  = SpreadsheetApp.getActiveSpreadsheet();
    var ss_ad  = sheet.getSheetByName(s_address);
    var ss_set = sheet.getSheetByName(s_setting);
    var ss_inf = sheet.getSheetByName(s_SetInfo);
    
    var row     = ss_ad.getLastRow();
    var address = ss_ad.getRange(2,1,row-1,2).getValues();
    var shop    = ss_set.getRange(s_shop_name[0],s_shop_name[1]).getValue();
    var boss    = ss_set.getRange(s_shop_boss[0],s_shop_boss[1]+1).getValue();
    var month   = ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]-1).getValue();
    
    // Mail Contents
    var title   = month+'月のシフト';
    var contents 
    = '.\n'
    + title + 'を作成しました.\n'
    + '以下のURLにアクセスし、全スタッフ'+lim +'日までに入力をお願いします.\n\n'
    + boss + '\n\n'
    +res;
    
    // Send Mail
    for(var i=0;i<address.length;i++){
      var mail = address[i][1];
      GmailApp.sendEmail(mail, title, contents);
    }
    
    Browser.msgBox('送信完了');
  }

}

// Send Complete Time Sheet
function NotifyStaffCompleteTS(){
  var res = Browser.msgBox("完成したシフトを送信しますか?", Browser.Buttons.OK_CANCEL);
  if(res === "ok"){
    var sheet  = SpreadsheetApp.getActiveSpreadsheet();
    var ss_ad  = sheet.getSheetByName(s_address);
    var ss_inf = sheet.getSheetByName(s_SetInfo);
    var ss1    = sheet.getSheetByName(s_setting);
    
    var row     = ss_ad.getLastRow();
    var address = ss_ad.getRange(2,1,row-1,2).getValues();
    var shop    = ss1.getRange(s_shop_name[0],s_shop_name[1]+1).getValue();
    var boss    = ss1.getRange(s_shop_boss[0],s_shop_boss[1]+1).getValue();
    var month   = ss_inf.getRange(s_timesheet_month[0],s_timesheet_month[1]-1).getValue();
    var pdfname = shop+'_'+month+'月_シフト.pdf';
    var file    = DriveApp.getFilesByName(pdfname).next();
    
    // Mail Contents
    var title   = month+'月のシフト';
    var contents 
    = '.\n'
    + title + 'が完成しました.\n'
    + '確認をお願いします.\n\n'
    + boss + '\n'
    
    // Send Mail
    Logger.log(address.length);
    for(var i=0;i<address.length;i++){
      var mail = address[i][1];
      GmailApp.sendEmail(mail, title, contents,{attachments: [file]});
    }
    
    Browser.msgBox('送信完了');
  }
}




////////////////////////////////////////////
////////// Update public holiday ///////////
////////////////////////////////////////////

// Get public holiday //
function GetPublicHoliday(url){
  var sh         = SpreadsheetApp.getActiveSpreadsheet();
  var ss_set     = sh.getSheetByName(s_setting);
  var ss_inf     = sh.getSheetByName(s_SetInfo);
  var SHEET_URL  = url;
  var SHEET_NAME = "祝日一覧";
  
  // From 1/1 in this year
  var startDate = new Date();
  ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]-1).setValue(startDate.getYear());
  startDate.setMonth(0, 1);
  startDate.setHours(0, 0, 0, 0);

  // To 12/31
  var endDate = new Date();
  endDate.setFullYear(endDate.getFullYear()+1, 11, 31);
  endDate.setHours(0, 0, 0, 0);  

  var sheet    = getholidaysheet(SHEET_URL,SHEET_NAME);
  var holidays = getHoliday(startDate, endDate);

  var lastRow  = sheet.getLastRow();
  var startRow = 1;

  // シートが空白で無いとき、取得した祝日配列の先頭の日付と一致するカラムの位置を探索
  if (lastRow > 1) {
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    for(var i = 0; i < lastRow; i++) {
      if(values[i][0].getTime() == holidays[0][0].getTime()) {
        break;
      }
      startRow++;
    }
  }

  sheet.getRange(startRow, 1, holidays.length, holidays[0].length).setValues(holidays);
  
}

function UpdatePublicHoliday(){
  var sh         = SpreadsheetApp.getActiveSpreadsheet();
  var ss_set     = sh.getSheetByName(s_setting);
  var ss_inf     = sh.getSheetByName(s_SetInfo);
  var url        = ss_inf.getRange(s_shurl[0],s_shurl[1]+1).getValue();
  var SHEET_URL  = url;
  var SHEET_NAME = "祝日一覧";
  
  // From 1/1 in this year
  var startDate = new Date();
  ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]-1).setValue(startDate.getYear());
  startDate.setMonth(0, 1);
  startDate.setHours(0, 0, 0, 0);

  // To 12/31
  var endDate = new Date();
  endDate.setFullYear(endDate.getFullYear()+1, 11, 31);
  endDate.setHours(0, 0, 0, 0);  

  var sheet    = getholidaysheet(SHEET_URL,SHEET_NAME);
  var holidays = getHoliday(startDate, endDate);

  var lastRow  = sheet.getLastRow();
  var startRow = 1;

  // シートが空白で無いとき、取得した祝日配列の先頭の日付と一致するカラムの位置を探索
  if (lastRow > 1) {
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    for(var i = 0; i < lastRow; i++) {
      if(values[i][0].getTime() == holidays[0][0].getTime()) {
        break;
      }
      startRow++;
    }
  }

  sheet.getRange(startRow, 1, holidays.length, holidays[0].length).setValues(holidays);
  
}


function getholidaysheet(SHEET_URL,SHEET_NAME){
  var ss    = SpreadsheetApp.openByUrl(SHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if(sheet == null) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  return sheet;
}

function getHoliday(startDate, endDate) {
  var cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");

  var holidays = cal.getEvents(startDate, endDate);
  var values   = [];

  for(var i = 0; i < holidays.length; i++) {
    values[i] = [holidays[i].getStartTime(), holidays[i].getTitle()];
  }

  return values;
}


////////////////////////////////////////////
//////////      All Function     ///////////
////////////////////////////////////////////

// Function1
function GetSetting() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s_setting);
  var ss_inf  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s_SetInfo);
  var ss_pday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s_hday);
  
  var DATE        = [];
  var Shop_time   = [];
  var Budget_time = [];
  var Rest_time   = [];
  var Pday        = [];
  
  // Day
  var ym               = ss_inf.getRange(s_timesheet_year[0],s_timesheet_year[1]-1,2,1).getValues();
  var date             = new Date(ym[0][0], ym[1][0]-1, 16);
  var ThisMonthLastDay = new Date(date.getYear(), date.getMonth()+1, 0); // 第3引数に0を入力する事でその月の最後の日を取得する事ができる．
  var NextMonth        = new Date(date.getYear(), date.getMonth()+1, 1);
  
  // Shop time
  var shop_time   = ss.getRange(s_opentime[0],s_opentime[1]+1,2,1).getValues();
  var budget_time = ss.getRange(s_workday[0],s_workday[1]+1,2,5).getValues();
  
  // Rest time
  var rest_time = ss.getRange(s_zero[0],s_zero[1]+1,4,2).getValues();
  
  // Public holiday
  var p_last = ss_pday.getLastRow();
  var y, m, d;
  for(var i=0;i<p_last;i++){
    now_date = Utilities.formatDate(new Date(ss_pday.getRange(i+1,1).getValue()), "JST", "yyyy/MM/dd");
    Pday.push(now_date);
  }
  
  DATE      = [date,ThisMonthLastDay, NextMonth];
  Shop_time = [shop_time,budget_time,rest_time];
  var ALL   = [DATE,Shop_time,Pday];
 
  return ALL; // returnで返すのは一つだけ.
  
}


// Function2
function DropDown(sh,s_r,s_c,f_r,f_c){
  var rule = SpreadsheetApp.newDataValidation();
  var list = ["×"
              ,"9:00","9:30","10:00","10:30","11:00","11:30","12:00","12:30"
              ,"13:00","13:30","14:00","14:30","15:00","15:30","16:00","16:30"
              ,"17:00","17:30","18:00","18:30","19:00","19:15","19:30","19:45"
              ,"20:00","20:15","20:30","20:45","21:00","21:15","21:30","21:45"
              ,"22:00","22:15","22:30","22:45","23:00"];
  rule.requireValueInList(list, true);
  sh.getRange(s_r,s_c,f_r,f_c).clearDataValidations();
  sh.getRange(s_r,s_c,f_r,f_c).setDataValidation(rule);  
}


// Function3 (Return timestamp)
function getTimestamp () {
  var now   = new Date();
  var year  = now.getYear();
  var month = now.getMonth() + 1;
  var day   = now.getDate();
  var hour  = now.getHours();
  var min   = now.getMinutes();
    
  return year + "_" + month + "_" + day + "_" + hour + min;
 }


// Function4 (Deleteformsheetbind)
function DeleteFormSheet(formshName,formid){
  var sh        = SpreadsheetApp.getActiveSpreadsheet();
  var form      = FormApp.openById(formid);
  var formsheet = sh.getSheetByName(formshName);
  form.removeDestination();
  sh.deleteSheet(formsheet);
  
}

// Function5 (CreateNewSpreadsheet)
function CreateSpreadsheetInfolder(folderId,fileName){
  var folder = DriveApp.getFolderById(folderId);
  var newSS  = SpreadsheetApp.create(fileName);
  
  var originalFile = DriveApp.getFileById(newSS.getId());
  var copiedFile   = originalFile.makeCopy(fileName, folder);
  DriveApp.getRootFolder().removeFile(originalFile);
  return copiedFile;
}  


function ChangeOwner(){
  try{
  var NewOwner = Browser.inputBox('ファイルを譲渡する相手のGMailアドレスを入力してください。');
  
  var SSId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(SSId);
  var OldOwner = Session.getActiveUser().getEmail();
  
  file.setOwner(NewOwner);
  file.removeEditor(OldOwner);
  
  Browser.msgBox("正常に譲渡されました。");
  }catch(e){
    Browser.msgBox("譲渡に失敗しました。\n再度、操作をお願いします。");
  }
}

function Copy(){
  var folderId = DriveApp.createFolder("シフト作成フォルダ").getId();
  var fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  var copyfilename = 'シフト作成フォーマット';
  var copy_dir = DriveApp.getFolderById(folderId);
  var moto_dat = DriveApp.getFileById(fileId);

  var newfile = moto_dat.makeCopy(copyfilename, copy_dir);
  
}
