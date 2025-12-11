//todo 入力形式を検討 (CSV が妥当？)

// 
// プロパティ指定子と動作内容対応
//
// Name
//     トラック名を設定する
//     使用する Property or Method : `HRESULT IITObject::Name  (  [in] BSTR  name   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITObject.html#z49_1</SDKREF>
// Artist
//     アーティスト名を設定する
//     使用する Property or Method : `HRESULT IITTrack::Artist  (  [in] BSTR  artist   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_5</SDKREF>
// Album
//     アルバム名を設定する
//     使用する Property or Method : `HRESULT IITTrack::Album  (  [in] BSTR  album   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_3</SDKREF>
// AlbumArtist
//     アルバムアーティスト名を設定する
//     使用する Property or Method : `HRESULT IITFileOrCDTrack::AlbumArtist  (  [in] BSTR  albumArtist   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_25</SDKREF>
// Year
//     年を設定する
//     使用する Property or Method : `HRESULT IITTrack::Year  (  [in] long  year   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_52</SDKREF>
// Compilation
//     コンピレーションかどうかを設定する true : 'コンピレーション' , false : 'コンピレーションではない'
//     使用する Property or Method : `HRESULT IITTrack::Compilation  (  [in] VARIANT_BOOL  shouldBeCompilation   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_12</SDKREF>
// DiscNumber
//     ディスク番号を設定する (1Base)
//     使用する Property or Method : `HRESULT IITTrack::DiscNumber  (  [in] long  discNumber   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_19</SDKREF>
// DiscCount
//     ディスク番号(全体)を設定する (1Base)
//     使用する Property or Method : `HRESULT IITTrack::DiscCount  (  [in] long  discCount   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_17</SDKREF>
// TrackNumber
//     トラック番号を設定する (1Base)
//     使用する Property or Method : `HRESULT IITTrack::TrackNumber  (  [in] long  trackNumber   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_48</SDKREF>
// TrackCount
//     トラック番号(全体)を設定する (1Base)
//     使用する Property or Method : `HRESULT IITTrack::TrackCount  (  [in] long  trackCount   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_46</SDKREF>
// AddArtworkFromFile
//     ファイルからアルバムアートワークを設定する
//     使用する Property or Method : `HRESULT IITTrack::AddArtworkFromFile  (  [in] BSTR  filePath,    [out, retval] IITArtwork **  iArtwork  )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z75_2</SDKREF>
// SortAlbum
//     アルバム名の '読み' を設定する
//     使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbum  (  [in] BSTR  album   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_39</SDKREF>
// SortAlbumArtist
//     アルバムアーティスト名の '読み' を設定する
//     使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbumArtist  (  [in] BSTR  albumArtist   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_41</SDKREF>
// SortArtist
//     アーティスト名の '読み' を設定する
//     使用する Property or Method : `HRESULT IITFileOrCDTrack::SortArtist  (  [in] BSTR  artist   )`
//                                  <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_43</SDKREF>

// プロパティ指定子チェック

// プレイリスト名として存在しない名称の文字列 `str_playlistName_tmp` を生成


// var	iTunesApp = WScript.CreateObject("iTunes.Application");

// 作業用プレイリストを作成
//todo 「作業用プレイリストを作成」する目的を記載
// var  IITPlaylist_playlist_tmp = iTunesApp.CreatePlaylist(filesystem.GetBaseName(str_tsvPath));

// 作業用プレイリストを削除

WScript.Echo("Done!");
