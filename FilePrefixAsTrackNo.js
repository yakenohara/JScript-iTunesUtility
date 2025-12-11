// iTunes で選択中のトラックのトラックNoとディスクNoをファイル名をもとに設定する

// <CAUTION>
// このファイルは文字コードを ※BOM付き※ UTF8 として保存すること。
// (`ü` などの文字化け回避)
// </CAUTION>

// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 


//iTunesObject生成
try{
	var itobj = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>
}catch(e){
	WScript.Echo("Cannot create object `iTunes.Application`");
	WScript.Quit(); // 終了
}

var fso = new ActiveXObject("Scripting.FileSystemObject");

//選択中のトラックの取得
var iittrackcol_tracks = itobj.SelectedTracks; // <SDKREF>iTunesCOM.chm::/interfaceIiTunes.html#z21_6</SDKREF>

if (!iittrackcol_tracks) { // トラックが選択されていない場合
    WScript.Echo("No tracks selected.");
    WScript.Quit(); // 終了
}

//<全体のディスク数・曲数を取得>------------------------------------------------------------------------------------------------

var intarr_maxes = []; // ディスク毎のトラックNo最大値を格納する配列 
                        // e.g. ディスク構成1枚かつ最大トラック数12の場合 -> [12]
                        // e.g. ディスク構成3枚かつ最大トラック数12, 13, 14の場合 -> [12, 13, 14]

var bl_isSingleDisc; // ディスク構成が1枚かどうか
var re_style_singleDisc = /^(\d+) /; // ディスク構成が1枚の時に使用する正規表現
var re_style_multiDisc = /^(\d+)\-(\d+) /; // ディスク構成が複数枚の時に使用する正規表現
var marr_matches_tmp;
var int_discNo_tmp;
var int_trackNo_tmp;

//Track 毎ループ
for( var int_idxOfTracks = 1 ; int_idxOfTracks <= iittrackcol_tracks.Count; int_idxOfTracks++ ){
    var objTrack = iittrackcol_tracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>

    // トラックのファイル名を取得
    var str_fileName = fso.GetFileName(objTrack.Location); //<SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_0</SDKREF>

    if (int_idxOfTracks == 1){ // 最初のトラックの場合

        // ディスク構成が1枚かどうかを判定
        marr_matches_tmp = re_style_singleDisc.exec(str_fileName);
        if(marr_matches_tmp !== null){ // ディスク構成が1枚の時のプレフィックスの場合 e.g.「01 trackname.m4a」 
            bl_isSingleDisc = true;
        }else{ // ディスク構成が1枚の時のプレフィックスでない場合
            bl_isSingleDisc = false; // ディスク構成は複数枚の時のプレフィックスを想定 e.g. 「1-01 trackname.m4a」
        }
    }
    
    if (bl_isSingleDisc){ // ディスク構成が1枚の場合
        marr_matches_tmp = re_style_singleDisc.exec(str_fileName);
        if (marr_matches_tmp === null){ // トラック No を指定するプレフィックスが見つからない場合
            WScript.Echo(
                "No prefix found to specify track number.\r\n" + 
                "Actual file name: " + str_fileName + "\r\n" +
                "Expected: \"01 trackname.ext\""
            );
            WScript.Quit(); // 終了
        }
        int_discNo_tmp = 1;
        int_trackNo_tmp = parseInt(marr_matches_tmp[1], 10);

    }else{  // ディスク構成が複数枚枚の場合
        marr_matches_tmp = re_style_multiDisc.exec(str_fileName);
        if (marr_matches_tmp === null){ // トラック No を指定するプレフィックスが見つからない場合
            WScript.Echo(
                "No prefix found to specify track number.\r\n" + 
                "Actual file name: " + str_fileName + "\r\n" +
                "Expected: \"1-01 trackname.ext\""
            );
            WScript.Quit(); // 終了
        }
        int_discNo_tmp = parseInt(marr_matches_tmp[1], 10);
        int_trackNo_tmp = parseInt(marr_matches_tmp[2], 10);
    }

    // ファイル名から抽出したディスクNoを反映
    if (intarr_maxes.length < int_discNo_tmp) { // 最大トラックNo格納先の要素が無い場合
        intarr_maxes[(int_discNo_tmp - 1)] = undefined; // 最大トラックNo格納先の要素を定義
    }

    // ファイル名から抽出したトラックNoを反映
    if (
        (intarr_maxes[(int_discNo_tmp - 1)] === undefined) ||
        (intarr_maxes[(int_discNo_tmp - 1)] < int_trackNo_tmp)
    ) {
        intarr_maxes[(int_discNo_tmp - 1)] = int_trackNo_tmp;
    }

}

//-----------------------------------------------------------------------------------------------</全体のディスク数・曲数を取得>

for( var int_idxOfTracks = 1 ; int_idxOfTracks <= iittrackcol_tracks.Count; int_idxOfTracks++ ){
    
    var objTrack = iittrackcol_tracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>

    // トラックのファイル名を取得
    var str_fileName = fso.GetFileName(objTrack.Location); //<SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_0</SDKREF>

    var int
    if (bl_isSingleDisc){ // ディスク構成が1枚の場合
        marr_matches_tmp = re_style_singleDisc.exec(str_fileName);
        int_discNo_tmp = 1;
        int_trackNo_tmp = parseInt(marr_matches_tmp[1], 10);

    }else{  // ディスク構成が複数枚枚の場合
        marr_matches_tmp = re_style_multiDisc.exec(str_fileName);
        int_discNo_tmp = parseInt(marr_matches_tmp[1], 10);
        int_trackNo_tmp = parseInt(marr_matches_tmp[2], 10);
    }

    objTrack.DiscNumber  = int_discNo_tmp; // <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_19</SDKREF>
    objTrack.DiscCount  = intarr_maxes.length; // <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_17</SDKREF>
    objTrack.TrackNumber = int_trackNo_tmp; // <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_48</SDKREF>
    objTrack.TrackCount  = intarr_maxes[(int_discNo_tmp - 1)]; // <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_46</SDKREF>
}

WScript.Echo("Done!");
