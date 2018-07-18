'30 november 2002 the stuff below
Private Declare Function GetVersionEx Lib "kernel32" Alias _
       "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias _
     "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
     lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
     hwndCallback As Long) As Long
'19 january 2003
'12 December 2004a
Private Declare Function mciGetErrorString Lib "winmm.dll" _
    Alias "mciGetErrorStringA" _
     (ByVal dwError As Long, _
     ByVal lpstrBuffer As String, _
     ByVal uLength As Long) As Long
'12 December 2004a

Private Declare Function getshortpathname Lib "kernel32" _
     Alias "GetShortPathNameA" _
     (ByVal lpszlongpath As String, _
     ByVal lpszshortpath As String, _
     ByVal cchBuffer As Long) As Long
'18 June 2003


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 '  Maintenance string for PSS usage
End Type

Private Type id3tag
    header As String * 3
    songtitle As String * 30
    artist As String * 30
    album As String * 30
    year As String * 4
    comments As String * 30
    genre As Byte
End Type
Private temptag As id3tag
'24 June 2003 above

' dwPlatforID Constants
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'-- End --'     30 november 2002

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Private rewind As String            '12Apr2014 use fastfwd as a rewind playback
Private copy_all As String          '15Dec2012
Private ser_num As String           '03Sep2011
Private cont_str As String          '08Feb2012
Private mixx As Integer             '30Jun2012
Private budgeyn As Integer          '17Dec2017 budge turned on by file name or at command promp #2
Private linebudgeyn As Integer          '17Dec2017 budge turned on a segment at a time
Private smlbud As Double            '17Dec2017  how much the video creeps ahead from the start position
Private bigbud As Double            '17Dec2017 When in fast forward this is the amount to jump ahead after but smlbud has played
Private maxbud As Integer           '17Dec2017 the max no of times the mini segment will play set to 30
Private totbud As Integer           '17Dec2017 the count on the way to maxbud
Private rand_prog As Integer        '29Oct2012
Private group_match As Integer      '29Oct2012
Private altt As Integer             '12Aug2012
Private mssg As String * 255        '26 February 2004
Private multi_prompt2 As String     '04 September 2004
Private hold_len As String          '20Feb2012
Private video_length As Long        '26 February 2004
Private slomo_point As Long        '25Feb2012
Private begin_locat As Long         '25Feb2012
'Private vs As String * 30            '06 February 2004
Private vs As String * 255            '19 March 2004  Having this the wrong size was a bug of major proportions...
Private last_vs As Double               '24 March 2004
Private replay_pos As Double            '10 April 2004
Private replay_yn As Integer            '10 April 2004
Private hold_speed As Double            '10 April 2004
Private command_line As String          '19 September 2004
Private back_job As String              '12 September 2004
Private testing As String               '10Feb2012
Private line_fit As String                       '03 August 2003
Private start_point As Double           '16 July 2003 Ver=1.07T
Private line_start_point As Double      '19 July 2003 Ver=1.07T
Private begin As String                 '15Sep2011
Private begin_point As Double           '20Sep2011
Private alt_amt As Double               '12Aug2012
Private resume_str As String            '25Sep2011
Private budge_str As String             '17Dec2017
Private again_str As String             '08Nov2011
Private keep_begin As String            '08Nov2011
Private keep_play_speed As Long       '08Nov2011
Private keep_line_delay_sec As Double        '08Nov2011
Private keep_resume_str As String       '08Nov2011
Private keep_line_start_point As Double      '08Nov2011
Private keep_begin_point As Double      '08Nov2011
Private keep_slomo As Integer           '08Nov2011
Private continueyn As String            '28Jan2012
Private pauseyn As String               '28Jan2012
Private hold_pauseyn As String          '25Mar2012
Private temp_double As Double           '19 July 2003 Ver=1.07T
Private thumb_nail As String            '16 July 2003 Ver=1.07T
Private elapse_start As Double          '13 July 2003
Private elapse_end As Double            '13 July 2003
Private elapse_yn As String             '13 July 2003
Private Pict_file As String  'february 09 2001 moved up to here for the "interrupt prompt"
Private long_pict_file As String    '22 November 2006
Private Save_file As String '15 January 2004
Private ooo As String       'february 09 2001 moved up here
Private xx1 As Integer      'february 09 2001 moved up to here for the "interrupt prompt"
Private yy1 As Integer
Private old_line As String  '25 March 2003 part of version ver=1.02b
Private motion_yn           '29 March 2003 part of version ver=1.04
Private special_date As String        '01 April 2003 reduce the delay if "Tuesday April 1" date display done as it has delays built in ver=1.05
Private time_displayed As String    '03 April 2003 indicate that time was displayed....
'Private replace_data As String      '03 April 2003 save SSS1 and replace it after date displayed
Private replace_sss1 As String
Private replace_sss2 As String
Private replace_sss3 As String
Private replace_sss4 As String
Private replace_sss5 As String
Private replace_sss6 As String      '03 April 2003 add the replace_sss? code above
Private sshortfile As String * 67   '18 june 2003
Private lresult As Long             '18 june 2003
Private videoyn As String   '10 february 2003
Private detailyn As String  '18 November 2004
Private avi_file As String  '01 february 2003
Private mpg_file As String  '11 May 2003    ver=1.05
Private wav_file As String  '10 June 2003   ver=1.07
Private mid_file As String  '10 June 2003   ver=1.07
Private auto_exe As String  '07 december 2002
Private ss_only As String   '07 december 2002
Private program_info        '21 december 2002
Private random_info         '21 december 2002
Private stretch_info        '21 december 2002
Private text_pause As Integer   '05 october 2002
Private debug_photo As Integer  '12 october 2002
Private show_files As String    '24 december 2002
Private show_files_yn As Integer '24 december 2002
Private do_tab As Integer  '05 october 2002
Private search_str As String    '26 august 2002
Private rand As Integer     '18 august 2002
Private rand1 As Integer    '23 March 2004
Private rand_cnt1 As Double  '23 March 2004
Private rand_no1 As Double  '23 March 2004
Private rand_cnt As Double  '18 august 2002
Private rand_no As Double  '18 august 2002
Private rand_str As String  '18Mar2012
Private diryes As String    'february 18 2002
Private Mergem As String    'february 24 2002
Private temptemp As String  'february 18 2002
Private tempdata As String  '31 March 2003
Private eofsw As String     '26 November 2004
Private fract_time As Double    'january 23 2002
Private check_time As Double    'january 23 2002
Private slomo As Integer        '08 January 2004
Private slomo_start As Double   '08 january 2004
Private motion_in As Double     '01 March 2004
Private motion_out As Double    '01 March 2004
Private slomo_in As Double      '08 January 2004
Private slomo_out As Double     '08 January 2004
Private pad_time As Double      '21 March 2004
Private slomo_seg As Double     '08 January 2004
Private inter_in As Double      '22 February 2004
Private inter_out As Double     '22 February 2004
Private sscreen_saver As String 'may 06 2001
Private sscreen_saver_ww As String '28 april 2002
Private os_ver As String        '30 november 2002
Private os_num As String        '01 december 2002
Private offset1 As Integer      'april 06 2001
Private offset2 As Integer      'april 06 2001
Private ppoffset1 As Integer    '27 july 2002
Private ppoffset2 As Integer    '27 july 2002
Private dblStart As Double      'used for elapsed time
Private dblEnd As Double        '  "   "      "     "
Private dbltime1 As Double
Private dbltime2 As Double
Private prompt2 As String   'march 01 2001
Private interrupt_prompt2   '20 November 2004
'see the 15 December 2004 minor changes so that "Y" for last of file can be defaulted in and out..
Private SAVE_ttt As String  'march 01 2001 controls defaults on start and default out after first entry
Private inin As String      'march 01 2001
Private SSS1 As String      'march 01 2001
Private prev_option As String   'february 25 2001
Private hh_cnt As Long         'april 22 2001
Private xtemp As String     'november 12 2000
Private ooopen As String    'september 02 2001
Private dsp_cnt As Long         'may 09 2001 count of total pictures displayed
Private delete_file As String 'june 30 2001
'insert the new globals right here october 28 2001
  Private strBuffer As String
  Private lngBufSize As Long
  Private lngStatus As Long
                Private lpBuff As String * 25
                Private ret As Long, UserName As String

Private strRootPathName As String
Private lngSectorsPerCluster As Long
Private lngBytesPerSector As Long
Private lngNumberOfFreeClusters As Long
Private lngTotalNumberOfClusters As Long

Private strDrive As String
Private strMessage As String
Private lngTotalBytes As Long
Private lngFreeBytes As Long
    Private hilite_hh As String     'april 22 2001
    Private hilite_cnt As Integer   'april 22 2001
    Private fs, f1, s 'april 11 2001
    Private dirdates As String  'april 11 2001
    Private indates  As String   'april 11 2001

    Private append_start1 As String  'april 10 2001
    Private append_end1 As String    'april 10 2001
    Private prev_ttt As String      'april 10 2001
    Private skipyesno As String     'june 09 2001
    Private old_pict As String      'april 08 2001
    Private newname As String       'april 01 2001
    Private filereason As String    'april 01 2001
    Private savepath As String      'april 01 2001
    Private leaddidg As String      'april 01 2001
    Private thedir As String        'april 01 2001
    Private indir As String         'april 01 2001
    Private outdir As String        'april 01 2001
    Private filetype As String      'april 01 2001
    Private dirtype As String       'february 23 2002
    Private minsize As Double      'april 01 2001 08 junt 2003
    Private maxsize As Double    'april 01 2001 08 june 2003 change from long to double
    Private autobuild As String      'april 01 2001
    Private inplace As String       'april 01 2001
    Private endstuff As String      'january 21 2001
    Private s1_imbed As String      'january 19 2001
    Private s2_imbed As String      'january 19 2001
    Private s3_imbed As String      'january 19 2001
    Private s4_imbed As String      '23 june 2002
    Private s5_imbed As String      '23 june 2002
    Private s6_imbed As String      '23 june 2002
    Private imbedded As String      'january 19 2001
    Private new_len As Integer      'january 18 2001
    Private array_ooo(55) As String 'january 18 2001
    Private array_aaa(55) As String 'january 18 2001
    Private match_flag As String    'january 18 2001
    Private array_pos As Integer    'january 18 2001
    Private array_prt As Integer    'january 18 2001
    Private data_ooo As String      'january 15 2001
    Private data_aaa As String      'january 15 2001
    Private over_lap As Integer     'january 10 2001
    Private crop_len As Integer     'january 09 2001
    Private wrap_cnt As Long        'january 07 2001
    Private mess_cnt As Long        'january 02 2001
    Private tot_s1 As Long          'january 01 2001
    Private tot_s2 As Long          'january 01 2001
    Private tot_s3 As Long          'january 01 2001
    Private tot_s4 As Long          '23 june 2002
    Private tot_s5 As Long          '23 june 2002
    Private tot_s6 As Long          '23 june 2002
    Private context_win As Integer  'january 01 2001
    Private page_prompt As String   'december 31 2000
    Private dateskip As String      'december 12 2000
    Private boundarycnt As Integer  'december 11 2000
    Private boundarystr As String   'december 11 2000
    Private mbxyes As String        'december 17 2000
    Private mbxi As Integer         'december 17 2000
    Private emailsea As String      'december 11 2000
    Private iimport As String       'december 24 2000
    Private showpos As String       'december 6 2000
    Private showasc As String       'december 11 2000
'    showpos = "Y"               'december 6 2000 **keep** very handy for testing line wraps only
    Private uppercase As String     'december 8 2000
    Private posstring As String     'december 6 2000
    Private crlf As String          'december 2 2000
    Private changes As String       'december 2 2000
    Private save_line As String    'to display error line numbers
                                'usefull for debug
    Private extract_yes As String   'november 12 2000
    Private skip_info As String     'november 14 2000
    Private encript As String       'november 20 2000
    Private hilite_this As String   'november 8 2000
    Private OutFile As Integer
    Private BatchFile As Integer    '27 October 2004
    Private FileFile As Integer     'march 20/00 edit files list
    Private CtrlFile As Integer     '19 November 2004
    Private NewFile As Integer      'january 29 2002
    Private ResultFile As Integer      '08 November 2004
    Private ExtFile      'november 12 2000
    Private FileExt As String   'november 17 2000
    Private case_yes    'november 13 2000
    Private search_prompt As String
    Private ampm As String
    Private hhour As String
    Private mminute As String
    Private ssecond As String
    Private ggg As String
    Private hhhour As Long
    Private mmminute As Long
    Private sssecond As Long
    Private chhour As Long
    Private cmminute As Long
    Private control_file As String 'november 03 2000
    Private control_files(10) As String '23 November 2004
    Private xxx_found As String    'november 03 2000
    Private pp_entered As String 'november 6 2000
    Private cssecond As Long
    Private freeze_sec As Double            '01 October 2003
    Private adjust_sec As Double            '29 September 2003
    Private delay_sec As Double
    Private fast_forward As Double          '06Feb2012
    Private new_delay_sec As Double         '05 March 2004
    Private line_delay_sec As Double        '19 July 2003 Ver=1.07T
    Private Line_freeze_sec As Double       '24 September 2003
    Private line_speed As Double            '16 November 2003
    Private hold_sec As Double  'february 08 2002
    Private photo_copy As String    'march 17 2001
    Private copy_photo As String    'march 17 2001
    Private photo_dir As String     'march 17 2001 ie c:\search\tempfold\
    Private bad_dir As String       '08Aug2016 mp3 with bad metadata ie pics go here d:\bad_mp3\
    Private photo_file As String    'march 17 2001 ie pict
    Private photo_cnt As Integer    'march 17 2001 sequential count of photo's
    Private temp_sec As Integer 'march 14 2001
    Private temp_cnt As Double  '08 July 2003
    Private temp_cnt1 As Double  '12 April 2004
    Private screen_capture As String    'march 15 2001
    Private tot_print As Integer
    Private ss_search As String
'may 06 2001    Dim sscreen_saver As String
'12 December 2004a
    Private result1 As Long          '12 December 2004a
    Private errormsg1 As Integer     '12 December 2004a
    Private returnstring1 As String * 1024   '12 December 2004a
    Private errorstring1 As String * 1024    '12 December 2004a
'12 December 2004a
    Private line_len As Integer
    Private entered_notes As String
    Private in_str As String
    Private out_str As String
    Private f As Integer
    Private FFF As String
    Private TheFile As String
    Private LastFile As String
    Private AllFiles(20) As String
    Private vvversion As String
    Private AllSearch(20) As String
    Private Cmd(100) As String
    'april 01 2001 made cript1 400 for the folders used in gf option
'    Private cript1(20000) As String   'november 18 2000
    Private cript1(25000) As String   '04 March 2010
'    Private cript2(100) As String   'november 18 2000
'    Private cript3(100) As Integer  'november 20 2000
'    Private cript2(20000) As String  '28 October 2004 big enough for my program...
'    Private cript3(20000) As Integer  '28 October 2004
    Private cript2(25000) As String  '04 March 2010
    Private cript3(25000) As Integer  '04 March 2010
'    Private criptcnt As Integer   'november 18 2000
    Private criptcnt As Long   '28 October 2004 above was 10000 too
    Private TheSearch As String
    Private FileFound As Integer
    Private Clip_data As String
    Private Last_match As String    'input line where match found
    Private last_pict           '28 february 2003
    Private noshow(10) As String
    Private screensave(10) As String
    Private screencount As Integer
    Private nocount As Integer
    Private Line_Search
    Private Picture_Search As String
    Private Previous_line_save As String
    Private Previous_line As String
    Private Next_line_save As String
    Private Next_line As String
    Private This_line As String
    Private String_Position As Integer
'    Dim Pict_file As String        'moved up for "interrupt prompt" february 09 2001
    Private disp_file As String
    Private long_line As String     'identify long line
    Private II As Integer
    Private EE As Double            '11 december 2002
    Private III As Integer
    Private JJ As Integer
    Private tt As Integer
    Private ttt1 As Integer
    Private ddd As Integer
    Private Enter_Count As Integer
    Private img_ctrl As String  'february 21 2001
    Private stretch_img As String   'march 31 2001
    Private auto_redraw As String   'november 10 2001
    Private ttt As String
    Private p2p2 As String      '10 august 2002
    Private tt1 As String   'testing only
    Private Test1_str As String 'Testing only
    Private Test2_str As String 'november 23 2001
    Private prompt1 As String
'    Dim prompt2 As String
    Private prompt3 As String
'    Dim SAVE_ttt As String
    Private qqq As String
 '   Dim SSS1 As String     what the hey why commented out????? ***vip*** it is strange too
    Private SSS2 As String
    Private SSS3 As String
    Private SSS4 As String  '09 june 2002
    Private SSS5 As String
    Private SSS6 As String
    Private sep As String
 '   Dim inin As String
    Private KEEPS1 As String
    Private KEEPS2 As String
    Private KEEPS3 As String
    Private KEEPS4 As String    '09 june 2002
    Private KEEPS5 As String
    Private KEEPS6 As String
    Private hi_lites As String          'flash display
    Private SAVE_KEEPS1 As String       'june 26/99
    Private SAVE_KEEPS2 As String
    Private SAVE_KEEPS3 As String
    Private SAVE_KEEPS4 As String   '09 june 2002
    Private SAVE_KEEPS5 As String
    Private SAVE_KEEPS6 As String
    Private Context As String           'june 26/99
    Private aaa As String
'    Dim ooo As String   'hold the original upper/lower case
    Private ccc As String   'aug 08/99
    Private xxx As String   'january 19 2001 checking for imbedded spaces in search string only
    Private Context_text(40) As String   'aug 08/99
    Private previous_picture(100) As Long
    Private pp As Long
    Private last_pp As Long 'may 10 2001
    Private skip_pp As Long 'may 10 2001
    Private previous_count As Long
    Private ddemo As String
    Private date_check As String
    Private today_date As String
    Private visual_impared As String
    Private Context_cnt As Integer  'aug 08/99
    Private Context_lines As Integer
    Private Clear_Context_lines As Integer  '18 March 2003 ver=1.01
    Private zzz_chrs As Long
    Private zzz_len As Integer
    Private cnt As Long
    Private SSS As String
    Private SAVE_SSS As String
    Private SAVE_SRCH As String
    Private MAX_CNT As Long
'    Dim back_cnt As Long
    Private bbb As Long
    Private tot_cnt As Long
    Private end_cnt As Integer
    Private tot_disp As Long
    Private time_cnt As Long
    Private time_num As Long
    Private lll As String
    Private tttpos As Integer
    Private i As Long
    
    Private loop_cnt As Long
    Private loop_inc As Long
    Private j As Long
    Private AltColor As Integer
    Private Hold_Fore As Integer        '27 July 2003
    Private Def_Fore As Integer
    Private Set_Fore As Integer
    Private temp_fore As Integer 'november 9 2000
    Private new1 As String 'november 9 2000
    Private new2 As String 'november 9 2000
    Private new3 As String 'november 9 2000
    Private new4 As String '09 june 2002
    Private new5 As String '09 june 2002
    Private new6 As String '09 june 2002
    Private mult1 As String 'november 21 2000
    Private mult2 As String 'november 21 2000
    Private mult3 As String 'november 21 2000
    Private mult4 As String '09 june 2002
    Private mult5 As String '09 june 2002
    Private mult6 As String '09 june 2002
    Private line_pos As String  'november 22 2000
    Private line_match As String 'november 21 2000
    Private keep_aaa As String 'november 9 2000
    Private keep_ooo As String 'november 9 2000
    Private temp1 As Long
    Private temp11 As Long      'march 14 2001
    Private temp2 As Long
    Private temp3 As Integer
    Private temp4 As Long       'november 9 2000
    Private temps As String
    Private testprompt As String    'january 24 2001
    Private ytemp As String     'december 11 2000
    Private badmp3_fold As String   '15Aug2016
    Private good_fold As String    '12Aug2016 the good mp3 directory
    Private tempss As String
    Private ppaste As String
    Private gpaste As String
    
    Private play_speed As Long   '29 October 2003
    Private save_play_speed As Long '16 November 2003
    Private date_displayed As String        'save for later search
    Private printed_cnt As Long             'may 10/00
    Private displayed_cnt As Long
    Private printed As String
    Private break_num As Long '27Aug2010
    Private grp As Integer      '15Sep2010 group number counter
    Private seq As Integer      '15Sep2010 sequence number count
    Private grpseq As String    '15Sep2010 display string in output file
    Private OffSet As Double      '02Nov2011 the offset value for time differences between various computers and devices when doing video
'october 28 2001
Private hold_zzz As Long    '21 September 2004
    Private yyy_cnt As Long         '28 October 2004
Private zzz_cnt As Long     'february 19 2001 display in interrupt prompt
'if items arn't declared in general they arn't available to the other subroutines.
'

Private Sub Command1_Click()
'    MsgBox xx1
    'may 06 2001 allow for an interrupt on the screen saver to switch into
    ' the p1 search and thus be able to do a "P" for previous picture
        Def_Fore = Hold_Fore  'reset color 27 July 2003
'  Test2_str = InputBox("27 July 2003 testing ", "testing doug ", , xx1 - 5000, yy1 - 5000)   '
'23 February 2004    If sscreen_saver = "Y" Then
'20 November 2004 with the code here the interrupt didnot work so put it back.. good test though
'20 November 2004    If sscreen_saver = "Y" And mpg_file <> "YES" And interrupt_prompt2 <> "WW" Then
'        frmproj2.Caption = "interrupt click " + ttt + "*" + Test1_str + "*" + prompt2 '20 November 2004
    interrupt_prompt2 = "WW"     '21 November 2004
    If sscreen_saver = "Y" And mpg_file <> "YES" Then
        ttt = "P"
        Test1_str = "P1"
        sscreen_saver = "N"
'  Test2_str = InputBox("delay_sec test ", CStr(new_delay_sec) + "*" + CStr(delay_sec), , xx1 - 5000, yy1 - 5000) '
'        frmproj2.Caption = "delay_sec test " + CStr(new_delay_sec) + "*" + CStr(delay_sec) '20 November 2004
        new_delay_sec = delay_sec   '23 November 2004 allow it to back up faster
        delay_sec = 0               '23 November 2004 delay_sec gets reset each new photo
'20 November 2004        SSS1 = "PHOTO"      'may 28 2001 allows the "P" to show all previous
                            'pictures after the screen saver is interrupted.
                            'it just looks at the last picture saved.
        prompt2 = "P1"
        GoTo little_lower       '21 November 2004
    End If          'may 06 2001
    If mpg_file = "YES" Then          '22 February 2004
    '02 March 2004 check the status here before pause
        DoEvents        '02 March 2004
        i = mciSendString("status video1 mode", mssg, 255, 0) '02 March 2004
        DoEvents        '05 March 2004
        frmproj2.Caption = " status_a is " + Trim(mssg) '05 March 2004
        DoEvents        '02 March 2004
        If Left(UCase(mssg), 6) = "PAUSED" Then
            i = mciSendString("resume video1", 0&, 0, 0)
            DoEvents
            mssg = "PLAYING"    '02 March 2004
        End If          '02 March 2004
        If Left(UCase(mssg), 7) = "PLAYING" Then
            i = mciSendString("pause video1", 0&, 0, 0)         '02 March 2004
            DoEvents
        Else
'Test2_str = InputBox("video status= " + Trim(mssg), "testing doug ", , xx1 - 5000, yy1 - 5000)
'            Print bel           '02 March 2004 testing only
'            frmproj2.Caption = " status is " + Trim(mssg)
        End If              '02 March 2004

'02 march 2004 i = mciSendString("pause video1", 0&, 0, 0) '22 February 2004
        inter_in = Timer    '22 February 2004 get the time we delay here.
        DoEvents        '22 February 2004
        i = mciSendString("status video1 position", vs, 255, 0) '22 February 2004
        tempdata = " **Pos=" + CStr(Val(vs)) + " (of) " + CStr(video_length)     '26 February 2004
        DoEvents        '23 February 2004
'26 February 2004        frmproj2.Caption = LCase(temptemp) + " pos=" + Trim(vs) '22 February 2004
        frmproj2.Caption = LCase(temptemp) + tempdata '26 February 2004
        DoEvents        '22 February 2004
    End If      '22 February 2004
    If mpg_file = "YES" Then            '24 February 2004
        If replay_yn = True Then        '10 April 2004
            play_speed = hold_speed     '10 April 2004
            replay_yn = False           '10 April 2004
        End If                          '10 April 2004  in case 2 interrupts in a row...
'27 August 2004      Test2_str = InputBox("Enter . to rewind video " + vbCrLf + Pict_file + vbCrLf + ooo, "Interrupt Prompt # " + CStr(dsp_cnt), , xx1, yy1) '24 February 2004
'02 September 2004    If thumb_nail = "YES" Then GoTo little_lower        '29 August 2004
'allow for thumbnail of video to pause first off maybe even do the same for mp3 (a pause is not bad here?)
        Test2_str = "A"      '02 September 2004
    If thumb_nail = "YES" And InStr(UCase(Line_Search), "MP3") <> 0 Then GoTo little_lower       '29 August 2004
'02 September 2004      Test2_str = InputBox("Enter . to rewind video (aa) for all" + CStr(video_length) + vbCrLf + Pict_file + vbCrLf + ooo, "Interrupt Prompt # " + CStr(dsp_cnt), , xx1, yy1) '24 February 2004
'08Nov2011 test the various saved values here
'25Dec2011 testing here doug
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
'      Test2_str = InputBox("E or X to Stop --  Enter . to rewind video (a) for all " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , xx1, yy1) '24 February 2004
'17Jan2012 default to "G" if mpegvideo here interrupt will cause the video clip to play again backspace to rid the "G"
        Test2_str = ""         '17Jan2012
        again_str = "NO"        '17Jan2012
        budge_str = "NO"        '17Dec2017
        budtot = 0              '17Dec2017
        If mpg_file = "YES" Then Test2_str = "G"        '08Feb2012 use for full rather than just again
        If cont_str <> "" Then Test2_str = cont_str '08Feb2012
'25Mar2012
'       Test2_str = InputBox("*0 pauseyn=" + pauseyn + " delay_sec=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + " E or X to Stop --  Enter . to rewind video (a) for all (F) or full (g) for aGain " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, Test2_str, xx1, yy1) '
'            If video_length = 0 And filetype = "MPG" Then
'      Test2_str = InputBox("*1 pauseyn=" + pauseyn + " start_point=" + CStr(start_point) + " delay_sec=" + CStr(delay_sec) + " E or X to Stop --  Enter . to rewind video (a) for all (F) or full (g) for aGain " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, Test2_str, xx1, yy1) '
'    If play_speed <> 1000 And play_speed <> 126 And again_str <> "YES" Then Test2_str = "" '22May2012 when speed set in file allow for interrupt to just stop the file
      Test2_str = InputBox(" E or X to Stop --  Enter . to rewind video (a) for all (F) or full (g) for aGain (s) for Slo-mo" + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, Test2_str, xx1, yy1) '
        If UCase(Test2_str) = "S" Then
           ' Test2_str = "G"
            cont_str = "H"      '16Apr2012 for switch between high and slow speede
            play_speed = 126           '14Apr2012
        End If                          '14Apr2012
        If UCase(Test2_str) = "H" Then
            'Test2_str = "G"
            cont_str = "S"      '16Apr2012 for switch between high and slow speede
            play_speed = 1000           '16Apr2012
        End If                          '16Apr2012
        If UCase(Test2_str) = "G" Or UCase(Test2_str) = "H" Or UCase(Test2_str) = "S" Then
'            If mpg_file = "YES" And hold_pauseyn <> "Y" And continueyn <> "Y" And again_str <> "YES" Then      17Dec2017 change add budge_str
             If mpg_file = "YES" And hold_pauseyn <> "Y" And continueyn <> "Y" And again_str <> "YES" And budge_str <> "YES" Then
'23Jan2018                delay_sec = (video_length - 200) / 1000
                delay_sec = video_length
                 line_delay_sec = delay_sec
                 keep_line_delay_sec = line_delay_sec
                 start_point = 100
                 line_start_point = start_point
                 keep_line_start_point = line_start_point
                 pauseyn = "Y"
            End If                      '25Mar2012
            again_str = "YES"       '27Nov2011
            budge_str = "YES"       '17Dec2017
 '15Dec2011           If begin_point < 11 Then  'something with this ***
 '15Dec2011               keep_begin = "YES"
                begin = keep_begin
'05Dec2011               begin_point = keep_line_start_point
                begin_point = keep_begin_point '05Dec2011
'05Dec2011                keep_begin_point = begin_point
               line_start_point = keep_line_start_point
'15Dec2011            End If
                ' GoTo line_15
'16Apr2012                cont_str = "G"      '08Feb2012
        End If                      '27Nov2011
        If UCase(Test2_str) = "F" Then
                again_str = "YES"       '08Feb2012
                begin = 10
                begin_point = 10 '08Feb2012
                delay_sec = 11111
                cont_str = "F"      '08Feb2012
                line_start_point = 10   '08Feb2012
                start_point = 10
 '               keep_begin = 10
 '               keep_line_start_point = 10
 '               keep_begin_point = 10
                         'testing
        End If                      '08Feb2012

'25Dec2011 testing here doug
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
'16Dec2011 Test2_str = InputBox("E or X to Stop --  Enter . to rewind video (a) for all " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , 4000, 5000) '24 February 2004
        Test2_str = Trim(UCase(Test2_str))        '12 December 2004
        If Test2_str = "E" Or Test2_str = "X" Then
            Unload Me       '12 December 2004
            Set frmproj2 = Nothing
            Stop                            '12 december 2004
'            Resume End_32000                   '12 December 2004  this line is not part of this subroutine...
        End If                              '12 December 2004 allow for exit too at interrupt...
little_lower:                                           '29 August 2004
'16 August 2004        frmproj2.Caption = LCase(temptemp) + " (Replay by Spectate Swamp)"  '26 February 2004
'        frmproj2.Caption = LCase(temptemp) + " (Replay by Spectate Swamp)length=" + CStr(video_length / 1000) + " seconds" '16 August 2004
        frmproj2.Caption = LCase(temptemp) + " " + rand_str + " (By Spectate Swamp)len=" + CStr(video_length / 1000) + " seconds" '16 August 2004
'02 September 2004    If thumb_nail = "YES" Then
    If thumb_nail = "YES" And UCase(Test2_str) = "A" Then
'
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
' Test2_str = InputBox("testing for thumb_nail= " + thumb_nail + " " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , 4000, 5000) '
'    If UCase(Test2_str) = "AA" Then
'            thumb_nail = "NO"
            Test2_str = ""
'            line_delay_sec = video_length / 1000
'            delay_sec = line_delay_sec
'            new_delay_sec = delay_sec
            i = mciSendString("stop video1", 0&, 0, 0)
            DoEvents
'            line_start_point = 10
'            i = mciSendString("play video1 from " + CStr(10) + " wait", 0&, 0&, 0&) '29 August 2004 this works better
            last_vs = video_length - 100    '29 August 2004 These 3 lines replace the WAIT call above works WELL
            slomo = True                '29 August 2004
            i = mciSendString("play video1 from " + CStr(10), 0&, 0&, 0&)
'      Test2_str = InputBox("When done hit enter " + vbCrLf + Pict_file + vbCrLf + ooo, "Waiting till end prompt # ", , xx1 - 5000, yy1 - 5000)  'february 09 2001
'            i = mciSendString("stop video1", 0&, 0, 0)
'            inter_in = Timer
'            GoSub line_30300
            GoTo down_abit
    End If              '27 August 2004
        DoEvents        '26 February 2004
    Else
      Test2_str = InputBox("Enter to continue x or e to exit " + vbCrLf + Pict_file + vbCrLf + ooo, "Interrupt Prompt # x" + CStr(dsp_cnt) + " " + save_line, , xx1 - 5000, yy1 - 5000) 'february 09 2001
    End If
    
'        ShowCursor = True        '29 november 2002
    Cmd(45) = ""        '07 december 2002 allow for other drives after interrupt..
    Test2_str = UCase(Test2_str)    '06 october 2002
'22 February 2004 on video play Pause, Show Position, Resume after adding in paused time.
    If mpg_file = "YES" Then          '22 February 2004
        DoEvents        '22 February 2004
'26 April 2004 try saving the results here
'        If Left(Test2_str, 1) = "F" Then
'            i = mciSendString("save video2 c:\search\outmpg.mpg", 0&, 0, 0)
'            new_delay_sec = 5.5
'            GoSub line_30300
'            Close video2
'            For EE = 1 To 10000000
'                DoEvents
'            Next EE                  'give it some time
'        End If
'need more info on the mcisendstring and how to record mpg files???
'26 April 2004
'24 february 2004 back up the video here if .... entered
'23 April 2004 allow for multiple ... or '''' to go back and ahead
'23 April 2004        If Left(Test2_str, 1) = "." Then
        If Left(Test2_str, 1) = "." Or Left(Test2_str, 1) = Chr(39) Then
            i = mciSendString("stop video1", 0&, 0, 0)
            III = Len(Test2_str)         '23 April 2004
            If Left(Test2_str, 1) = Chr(39) Then III = III * -1     '23 April 2004
            DoEvents
'10 April 2004 allow for replay at a slower speed dictated by a cmd(?) value...
            replay_yn = False               '10 April 2004
            i = mciSendString("status video1 position", vs, 255, 0) '10 April 2004
            temp3 = InStr(vs, Chr$(0)) '10 April 2004
            temp11 = Val(Left(vs, temp3 - 1)) '10 April 2004
'23 April 2004           If Val(Cmd(66)) > 1 And Val(Cmd(66)) < 1000 Then
'23 April 2004  do not change speed if moving ahead in file
'08Nov2011?? check the cmd(66) element for problems
           If Val(Cmd(66)) > 1 And Val(Cmd(66)) < 1000 And Left(Test2_str, 1) <> Chr(39) Then
                hold_speed = play_speed     '10 April 2004
                play_speed = Val(Cmd(66))   '10 April 2004
                replay_yn = True            '10 April 2004
                replay_pos = temp11        '10 April 2004
                slomo = True                '10 April 2004  if full speed to slow speed this needed
            End If                          '10 April 2004
'23 April 2004            i = mciSendString("play video1 from " + CStr(Val(vs) - (Val(Cmd(27)) * 1000)), 0&, 0&, 0&)
            i = mciSendString("play video1 from " + CStr(Val(vs) - (Val(Cmd(27)) * 1000) * III), 0&, 0&, 0&)
            last_vs = Val(vs) - Val(Cmd(27) * 1000)   '06 April 2004 without this the replay is at fast speed
                                    'might be handy to have this as an option maybee
            DoEvents
'        temptemp = InputBox("04 April 2004 test ", "Test continue " + CStr(dsp_cnt), , xx1 - 5000, yy1 - 5000)
            Test2_str = ""
'testtest = interrupt_prompt2 + " " + resume_str + " " + motion_yn + " " + CStr(line_start_point) + " " + CStr(start_point) + " " + CStr(begin_point) + " " + CStr(play_speed) + " " + CStr(slomo) + " " + CStr(delay_sec) + " " '08Nov2011 teststring again info
'        testtest = InputBox("again info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            GoTo down_abit
        End If                      '24 February 2004
        i = mciSendString("resume video1", 0&, 0, 0)
        DoEvents        '09 March 2004
down_abit:      '24 February 2004
        inter_out = Timer   '22 February 2004 add the delay here back in.
        If slomo = False Then
            fract_time = fract_time + inter_out - inter_in  '22 February 2004
        End If
    End If      '22 February 2004
'22 February 2004
    
'06 January 2005 allow for an entry to continue here
'06 January 2005    If text_pause And Test2_str <> "C" Then
    If text_pause And (Test2_str = "X" Or Test2_str = "E") Then
        text_pause = 0      '05 october 2002    "C" allow for continue of auto display
        inin = ""
    End If
    If UCase(Test2_str) = "A" Then
        prompt2 = "C"        'march 01 2001
        SAVE_ttt = "C"
        inin = "A"
        SSS1 = "A"      'march 01 2001
    End If
'13 March 2004  If UCase(Test2_str) <> "X" And UCase(Test2_str) <> "E" Then GoTo line_15  'February 09 2001
    If mpg_file = "YES" Then
'04 April 2004        i = mciSendString("close all", 0&, 0, 0)
        DoEvents
'22 March 2004        tt1 = InputBox("testing only closing down", , , 4400, 4500)  'TESTING ONLY 13 March 2004
    End If                      '11 March 2004
'04 April 2004        Unload Me
'04 April 2004    Set frmproj2 = Nothing 'sdistuff
'04 April 2004           Set colReminderPages = Nothing  'release memory??
        
'    Set sear1 = Nothing  'mdistuff
'04 April 2004 take this out    End             'feburary 5 2001
line_15:
'        temptemp = InputBox(" 04 April 2004 doug  " + SSS1, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
End Sub

Private Sub Form_Click()
'this is activated in the flash mode when mouse button is hit march 02 2001
  Test1_str = InputBox("in form_click" + vbCrLf + Pict_file + vbCrLf + ooo, "Interrupt Prompt " + CStr(zzz_cnt) + " " + save_line, , xx1 - offset1, yy1 - offset2) 'february 09 2001
'march 02 2001
'seems not to do anything when in sdi mode ???
'ie holding down the left mouse does nothing only works in mdi mode
End Sub

Private Sub Form_Deactivate()
'03 July 2003 testing to clean up the unload process (just needed to add these two routines?????)
'using the X for closing the form down check this bit out
End Sub

Private Sub Form_Initialize()
'system compoenents:    Use
'form                   does almost everything
'command button         used for esc out of screen saver
'                       set cancel property to true and place off the screen
'text block (removed)   allows key entry to interrupt screen saver (placed off the screen)
'image control          allows for p2 option with stretch set to true for full screen size
'when compiling for the millennium set BorderStyle to 2 from 0 todo **vip**
    'Load proj2
    ' Always set the working directory to the directory containing the application.
' do the startup steps below
'        tt1 = InputBox("doug startup test1 " + App.Path, "titlename", "Default", 4400, 4500)  'TESTING ONLY
    
    ChDir App.Path      'december 1 2000
                        'same as "set default" on the vax
'    frmproj2.Caption = program_info + " (cls #20)" '18Dec2013
    Cls         'november 7 2000
'january 31 2001
'App.StartLogging "douglog.txt", 0    'july 04 2001 see other test of startlogging
'the above did create a file so this is a start july 04 2001

'january 31 2001
'31July2011 comment out the following 6 lines for already running 29Jan2012 re-activate it
   If App.PrevInstance Then
      MsgBox "Already running!", _
         vbOKOnly, App.EXEName & _
         "Warning!"
      End
   End If
offset1 = 4500      'april 06 2001
offset2 = 5000      'april 06 2001
    program_info = "Spectate Swamp Content Management App"   '17Mar2016
    App.Title = program_info      'january 31 2001
    frmproj2.Caption = program_info + " (stonedan@telusplanet.net)" '25 november 2002
    frmproj2.Show   'sdistuff
'    sear1.Show      'mdistuff
    'equivalent of CVT$% is Asc below
 '        tt1 = InputBox("Form_Initialize " + CStr(Asc("a")) + Chr(97), , , 4400, 4500) 'TESTING ONLY
'        tt1 = InputBox("Form_Initialize " + App.Path, "titlename", "Default", 4400, 4500)  'TESTING ONLY
  '  Cls
    
'february 05 2001 comment out the next line
'    Call text2_Change  ' february 01 2001
    Call text2_Chg    'february 05 2001
    'test the logging stuff some time later....
'    Call App.StartLogging(App.Path + "\logfile.txt", 2)    'january 30 2001
'    Call App.LogEvent("starting", 4)       'january 30 2001
'    Call App.LogEvent(txtLog, vbLogEventTypeInformation) 'january 30 2001
End Sub

Private Sub text2_Chg()
'Private Sub text2_Change   'february 01 2001
    'savior version 1.0
    'all unpaid copies carry a curse
    '   be warned
    '
    ' ** set auto_redraw to "false"
    '       much much faster than true
    '       for Keep-it super-search anyway
    '    set windowstate to 2 for maximized
    'to get rid of the mini text block set the
    '   left element from 402 to 600
    'in the properties window for the for set the following
    'height to 9000
    'width to 12000
'==========================================================================
'january 31 2001   see function above GetComputerName

  
  lngBufSize = 255
  strBuffer = String$(lngBufSize, " ")
  lngStatus = GetComputerName(strBuffer, lngBufSize)
  strBuffer = Left(strBuffer, lngBufSize)
  If lngStatus <> 0 Then
'commented out january 31 2001 below
'     MsgBox ("Computer name is: " & Left(strBuffer, lngBufSize))
  End If
'==========================================================================
'Code:          see function above GetUserName
     ' Main routine to Dimension variables, retrieve user name
     ' and display answer

                ' Dimension variables

                ' Get the user name minus any trailing spaces found in the name.
                ret = GetUserName(lpBuff, 25)
                UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

                ' Display the User Name
'                MsgBox UserName        'commeted out january 31 2001
'==========================================================================
'==========================================================================
'see call to GetDiskFreeSpace function above january 31 2001

'Code:
    strDrive = "C:\" 'drive letter
    
'note to vb6 comment out the following code as it gets errors overflow etc...
'january 27 2002 comment it out as I am not using it anyway
'    If GetDiskFreeSpace(strDrive, lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters) = 0 Then
'        strMessage = strMessage & vbCrLf & "An error occurred."
'    Else
'        strMessage = strMessage & vbCrLf & "Sectors Per Cluster: " & Format$(lngSectorsPerCluster)
'        strMessage = strMessage & vbCrLf & "Bytes Per Sector: " & Format$(lngBytesPerSector)
'        strMessage = strMessage & vbCrLf & "Free Clusters: " & Format$(lngNumberOfFreeClusters)
'        strMessage = strMessage & vbCrLf & "Total Clusters: " & Format$(lngTotalNumberOfClusters)
'        lngTotalBytes = lngTotalNumberOfClusters * lngSectorsPerCluster * lngBytesPerSector
'        strMessage = strMessage & vbCrLf & "Total Bytes: " & Format$(lngTotalBytes)
'        lngFreeBytes = lngNumberOfFreeClusters * lngSectorsPerCluster * lngBytesPerSector
'        strMessage = strMessage & vbCrLf & "Bytes Free: " & Format$(lngFreeBytes)
'        strMessage = strMessage & vbCrLf & "Percent Used: " & Format$(1 - (lngFreeBytes / lngTotalBytes), "0.00%")
'    End If
    
    If Left(strBuffer, lngBufSize) = "OEMCOMPUTER" Then
'        MsgBox (strMessage)
    End If                          'january 31 2001 only show it if it my laptop
    strMessage = strMessage + vbCrLf + "User name=" + UserName + vbCrLf + "Computer name=" + strBuffer & vbCrLf

'30 november 2002 the stuff below for computer version
   Dim tOSVer As OSVERSIONINFO
    
   ' First set length of OSVERSIONINFO
   ' structure size
   tOSVer.dwOSVersionInfoSize = Len(tOSVer)
   ' Get version information
   GetVersionEx tOSVer
   ' Determine OS type
   With tOSVer
      
      Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32_NT
            ' This is an NT version (NT/2000)
            ' If dwMajorVersion >= 5 then
            ' the OS is Win2000
            If .dwMajorVersion >= 5 Then
               FFF = "Windows 2000"
            Else
               FFF = "Windows NT"
            End If
         Case Else
'  Test2_str = InputBox("testing versioninfo " + CStr(.dwPlatformId) + " " + FFF, " " + CStr(zzz_cnt), , xx1, yy1) 'february 09 2001
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
               FFF = "Windows ME"
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
               FFF = "Windows 98"
            Else
               FFF = "Windows 95"
            End If
         End Select
         ' Check for service pack
         FFF = FFF & " " & Left(.szCSDVersion, _
                          InStr(1, .szCSDVersion, Chr$(0)))
        os_ver = FFF        '30 november 2002 I use this variable later maybe
         ' Get OS version
         ttt = "Version= " & .dwMajorVersion & "." & _
                          .dwMinorVersion & "." & .dwBuildNumber
    os_num = .dwMajorVersion & .dwPlatformId      '01 december 2002 use to determine 8, 10 or 43, 44
    End With
strMessage = strMessage + ttt + " len=" + CStr(Len(os_ver)) + vbCrLf
strMessage = strMessage + "OS=" + os_ver  'I used this for display
'        tt1 = InputBox(strMessage, "testing", , 8000, 5000)

'30 november 2002
'==========================================================================


    On Error GoTo Errors_31000
    save_line = "000"           'for error handling
'    Static xx1 As Integer
'    Static yy1 As Integer

        '0 = black
        '1 = dark blue
        '2 = green
        '3 = aqua
        '4 = brown
        '5 = purple
        '6 = lime/green/brown?
        '7 = grey
        '8 = dark grey
        '9 = blue bright
        '10 = lime green bright
        '11 = pale blue
        '12 = red
        '13 = purple light pink
        '14 = yellow
        '15 = white
'test some code here in this area
'temps = CallByName(wordpad, "c:\search\notes.txt", , "c:\search\notes.txt")
line_20:
    save_line = "20"
'04 November 2003 see cmd(61) for play_speed    play_speed = 1000           '30 October set to normal speed
    SAVE_ttt = "in"             'default to "c" if comming in
                                'april 14/00
    start_point = 0             '16 July 2003 Ver=1.07T
    start_point = 10            '09 September 2003 service pack 1 needs this ????
    line_start_point = 0        '19 July 2003 Ver=1.07T
    line_start_point = 10       '11 March 2007
    search_prompt = "in"        'april 24/00
    visual_impared = "YES"      'to set up control.txt
    visual_impared = "NO"       'settings for ron windels
    date_check = "NO"       'do not do date check logic
'    date_check = "YES"      'do date check logic 26Feb2012
    
    ddemo = "YES"
    ddemo = "NO"    'set to "YES" if demo "NO" if full system
    If UCase(App.EXEName) = "DEMO" Then ddemo = "YES"   '08 August 2003
'    control_file = "c:\control.txt" 'november 03 2000
    control_file = "control.txt" 'december 03 2000
    GoSub Control_28000     'assign font color etc 20/04/00  initial read control file
'19Aug2011 change the control file here if app.exename is in cmd(82)
 If InStr(1, UCase(Cmd(82)), UCase(App.EXEName)) <> 0 Then
    control_file = "control1.txt" '19Aug2011
    GoSub Control_28000
 End If         '19Aug2011
'    Picture1.Left = 0       'june 17/99
 '   Picture1.Top = 0
    save_line = "25"
   OutFile = FreeFile
    KEEPS1 = ""
    KEEPS2 = ""
    KEEPS3 = ""
    KEEPS4 = ""     '09 june 2002
    KEEPS5 = ""
    KEEPS6 = ""
    ser_num = "0000001"     '03Sep2011   some of the initialization here
        'tell them where this program comes from
        SSS = Format(Now, "ddddd ttttt")    'todays date
        temp1 = InStr(2, SSS, "/1/")
'        vvversion = "Version September 1/00"
'        vvversion = "Version September 5/00"  'extended control.txt to 40 from 20
'        vvversion = "Version September 8/00"  'line wrap control.txt element 21 to 40 from 82 etc
'        vvversion = "Version September 10/00" 'context lines 10 or 5 etc
                                       'added the B for back option
'        vvversion = "Version October 4/00" 'noshow and others added
'        vvversion = "Version October 9/00" 'p for previous picture allowed
'        vvversion = "Version October 13/00" 'restricted demo version and date check back in
'        vvversion = "Version October 19/00" 'don't load the clipboard and make sure search selection has at least 3 of the same chrs
'        vvversion = "Version October 21/00" 'Quickie search if Q entered no context display and skip if no match
'        vvversion = "version October 23/00" 'quickie fixes and security completed next audio
'        vvversion = "Version October 26/00" 'no start character required and unload form3 at end automatically no close required
'        vvversion = "Version October 27/00" 'allow for "A" for all at next photo prompt
'        vvversion = "Version October 29/00" 'finish with line wrap fixes context data and match line fixed finally
'        vvversion = "Version November 3/00" 'allow for ccc to select control1.txt for larger fonts etc
                                            'allow for auto setting of c and p1 depending if xxx. found in first 10 lines
                                            'allow for any file selection to reset as if first time in for defaults etc
'        vvversion = "Version November 12/00" 'allow for extract of printed data to file
'        vvversion = "Version November 13/00" 'make rrr case sensitive and allow or name as of files
'        vvversion = "Version November 14/00" 'allow for "F" at do you want to continue
                                             'as well change the xxx to be entered at the option entry
'        vvversion = "Version November 22/00" 'changes to allow for multiples on a line to be displayed ie telusplanet.net
                                             ' both the net's will show on a "net" search
'        vvversion = "Version November 29/00" 'minor bug fix to the rrr routine
'        vvversion = "Version December 2/00"  'allow for QQQ at the option level same as RRR but
                                             'used to fix From: Subject: in mail messages removes crlf at end
                                             'and pastes the next line on to the end Date: Organization: To: done too
'        vvversion = "Version December 3/00"  'use default directory App.Path wherever the program exists look for control.txt etc.
'        vvversion = "Version December 5/00"  'do display of multiples on a line minor change
'        vvversion = "Version December 6/00"  'allow for LL linelength to be input and for showpos options
                                             'the showpos shows the line length at the wrap just in case more problems crop up
'        vvversion = "Version December 7/00"  'allow for Cmd(33) to be search default ie "D" for me and use the inputbox prompt field
                                             'do the changes to get rid of the max and min border set borderstyle to 0 from 2 and max and min to false
'        vvversion = "Version December 8/00"  'change option S to work like option Q quickie search but not to be case sensitive. it takes 2 as long as Q not 4 times as long now
        'allow for X or E to exit if element 30 set to "N" and make files case insensitive?
'        vvversion = "Version December 10/00" 'reget defaults once file is selected a second time
'        vvversion = "Version December 11/00" 'Allow for showasc at the option level to show the last 4 characters ascii value.
                                             'Allow for email prompt at the option level to read thru and display email records without the junk
'        vvversion = "Version December 20/00" 'minor changes to fix the netscape extract etc
'        vvversion = "Version December 22/00" 'allow for "A" for all when using the "Q" search
'        vvversion = "Version December 23/00" 'a few fix ups for the imported outlook express into netscape 6.0
'        vvversion = "Version December 24/00" 'new option of "IMPORT" if the file is email and import entered
'        vvversion = "Version December 27/00" 'if CH was selected the option prompt was getting in the way the second time around
                                             'the AppActivate fixed the problem though
'        vvversion = "Version December 28/00" 'did some clean up on the search selection no longer allow the number of characters to select previous searches
                                             'minor fix to the "." at the b for back prompt cleared SSS2 and SSS3 etc
'        vvversion = "Version December 31/00" 'minor fix re short lines with odd characters and boundarystr reset to ""2
                                             'new option "R" extract skip page_prompt AND set to "Q" search
                                             'as long as everything is being extracted then don't stop at number of lines per screen Cmd(1)
                                             'new option "D" extract skip page_prompt and set to "C" with MAX_CNT = 5 Context_lines = 2
'        vvversion = "Version January 01/01"  'new option "T" just like D and R above
                                             'use element Cmd(34) for the context_win default to 3 if not there
'        vvversion = "Version January 02/01"  'final counts for tot_s1 tot_s2 tot_s3 now tested and seem to work ok
'        vvversion = "Version January 03/01"  'noshow for text display will replace the string with **** as censorship
                                             'the 03a version changes made to allow elements to be hilited after the "A" is entered.
                                             'the very first one on a line was what was happening and this was the fix
'        vvversion = "Version January 04/01"  'fix for multiple no-show elements and a 04a fix to show the last elements
                                             'which were sometimes missed.
                                             '04b changes to allow for the "B" back to retain the hiliting capabilities
'        vvversion = "Version January 05/01"  'remove reference to Cmd(13) back_cnt and Cmd(19) for re-use
                                             'fixed the "D" date display by searching for "M" as both AM and PM have it..
                                             'if "." default to last search any other chr back to beginning and new search entry.
'        vvversion = "Version January 06/01"  'skip duplicate code at line_12004b thru line_12005m as not required, moved some noshow logic
                                             'do not include the counts for hilite_this or "append start" or "append end"
'        vvversion = "Version January 08/01"  'fix the logic for wrap_cnt somewhat so the V & B for back work a little better (still not completely checked out)
'        vvversion = "Version January 09/01"  'add the "crop" option to shorten very long lines that cause wraping problems
'        vvversion = "Version January 10/01"  'allow for the over_lap element 13 on the crop and the wrap function later.
'        vvversion = "Version January 19/01"  'the wrap array logic mostly complete with code bypassing this till the 12000 routine completed.
'        vvversion = "Version January 20/01"  'the new wrap array and hi-lite logic testing being done
'        vvversion = "Version January 21/01"  'add in logic to print last "endstuff" if A for All and wrap info unprinted
'        vvversion = "Version January 22/01"  'removed the old hi-lite wrap logic at sub_12000:
'        vvversion = "Version January 24/01"  'minor fix to hi-lite elements that wrap on to next page...
'        vvversion = "Jan 25/01"  'change the A for append for Z to paste at the end of the file (confused with a for all)
'        vvversion = "Jan 27/01"  'fix to skip the noshow elements if SS screen saver
'        vvversion = "Jan 28/01"  'use tot_s1 to count the number of emails on the import function
'        vvversion = "Jan 31/01"  'check if app is running and set App.Title to Keep-It
'        vvversion = "Feb 01/01"  'check for computer name mine and ken's and anyone else
'        vvversion = "Feb 06/01"  'allow the esc to stop the screen saver leave in old logic for any other key though
'        vvversion = "Feb 09/01"  'if interrupted allow for them to continue
'        vvversion = "Feb 16/01"  'add the date and time on the file prompt line
'        vvversion = "Mar 01/01"  'allow for "A" for all to be entered at the interrupt prompt
    'Feb 19/01 display the file name in each of the prompt boxes
    'feb 20/01 setfocus to solve the problem of the interrupt not working after "CH" called then F or SS done
    'feb 21/01 add image control set the stretch property to true (for full screen display)
    'feb 25/01 add check for only "CH" when using setfocus as it seems to screw up other functions?
    'feb 28/01 check for bounary=""" as it crapped out when wrong boundary found
    'Mar 14/01 allow for WAIT= to pause line longer ie wait 30 etc fix the timer error at line_30000: area
    'Mar 15/01 allow for SC screen capture to pause before the inputbox statements so that the displayed info can be capured using alt/print screen
    'Mar 17/01 allow for CP copy picture to a directory in Cmd(23) element
    'Mar 21/01 add in a was=scn0321.jpg to the text on a CP so we know the original number & date
    'Mar 31/01 set the P1 as the stretch default and P2 non stretched (they will have to user P2 for sizing original scans)
    'Apr 01/01 the GF get file routine to catalogue any pictures in the target directory
    'Apr 05/01 allow for keeping the same name if no leading didgits entered.
    'Apr 06/01 allow for 5000 folders when using the gf option
    'Apr 08/01 allow for c:\search\tempfold\ to keep the original name on file copy
    'Apr 10/01 allow for HL and HLL to use the append_start1 and append_end1 elements hilite_this could also be used???
    'allow for uppercase on all 3 of them working well so far
    'Apr 11/01 allow for c:\* to do the whole c drive and * FOR file type
    'and put out the date if dirdates entered
    'Apr 22/01 allow for hhxxx. to display pictures at the end of screen promopt
    ' by putting in a HH when some hilited.
    'may 06 2001 on interrupt of screen saver switch into photo display p1
    ' allowing the user to do a "P" to see one of the various previous pictures
    ' as well at the photo continue prompt allow "SS" to continue screen saver
    ' with the previous screen saver settings
     '   vvversion = "May 06/01"  'this is the date of the compile only, just for the users info
    'vvversion = 28 april 2002   add in the ww like ss with anded search criteria
'   the Nov2011 changes were for the BEGIN AND AGAIN options which were sweet
        'ver=1.04       added the ability to clear date to left by removing 1 character at a time and
        '               re-displaying what is left (looks like motion)
'28Jan2012 4.90 version changes were to allow for the continue and pause statements
'with the 28Jan2012 changes make sure that if using multiples for a wait that the file is "CLOSED" ???
'31Jan2012 the cmd(48) setting to photo required to make the continue work with the pause... change the default on create
'07Feb2012 put in a minor check to catch when video_length is obviously WRONG
'09Feb2012 check for mpg_file when setting delay second needed to be done
'09Feb2012 add "NORG" and "RG" for cmd(81) noRAND_GROUP need to switch back and forth for ff cataloged stuff
'10Feb2012 testing to see why the fast forword would work on the pc and not the laptop
'          left the debug lines in with this date it seems if the time for play segment is larger it works???
'17Mar2012 put a trim on the search for BREAK the end of file was looping....
'18Mar2012 display the random number it seems to be a off some times when there is not much data
'25Mar2012 make changes so the enter enter caused a replay to happen
'14Apr2012 change so the error opening video file message does not show error it scares people
'14Apr2012 and allow for a replay to be at slow motion  minor bug with random fixed as welllllll
'16Apr2012 allow to switch between slow motion and full speed by hitting enter enter
'12May2012 force the prompts for RANDOM and SEQUENTIAL by name as well as control file
'14May2012 fix a few problems with the 12May2012 and one other small bug
'18Jun2012 allow for program name to control logic flow see catmydrive and a few other changes
'22Jun2012 bigtext changes to allow for automatic large print using bigtext.txt as input file
'23Jun2012 catmydrivemp3 added and minor changes to bigtext
'23Jun2012a allow for auto run of randomseg to play thumbnails with random video and random start point
'26Jun2012 minor fix to cmd(45) so MANUAL runs with prompts
'30Jun2012 through 29Jul2012 changes to allow for mixx of text and video
'07Aug2012 default to no update for every one beyone D
'12Aug2012 allow for alt== in the line to change the start== and begin== points re the pc vs laptop weird start points
'08Sep2012 change so that next pass on mixx stuff would show videos bug fixed but underline persists no big deal
'15Nov2012 change to allow for random mix stuff any changes between 08sep and 15Nov may cause some trouble so beware
'15Dec2012 allow for doall when doing a copy picture with lots of files
'17Mar2014 changes to allow for fastfwd catalog for video collection
'25Mar2014 minor change so rand_group works
'16Apr2014 new feature bb.08 option #2 prompt does the reverse of ff.08 fast forward reverse option for switch
'          needs to run reverse right after to switch the records completely
'17Mar2016  Version on Windows 10 laptop
'09jun2016        vvversion = "17 March 2016"  this is the date of the compile only, just for the users info
'07Sep2016  mp3 bad records to a 2nd output file the file size on directory command etc
'30Sep2016 add the "COLOR=" option to switch to aqua etc and overprint on a bitmap screen capture of same text etc wipe it clean
        vvversion = "version W10.2 12Mar2018"     '12Mar2018 Windows 10 latest budge feature
'look for "check the latest features above" where to add new features to help file

        If temp1 <> 0 And ddemo = "YES" Then
        Print vvversion; " of " + program_info
 '      Print "check web site: http://www.telusplanet.net/public/stonedan"
        Print "e-mail Doug Pederson stonedan@telusplanet.net"
        Print "for newer versions with more options"
        Print "shareware"
        Print "this info only displays on the first of month"
        Print "Enter to continue"
        tt1 = InputBox("Enter to continue", "Continue Prompt", , xx1 - offset1, yy1 - offset2)
        End If              'year must be 2000

'   do not allow use if date past best before date???
'        Print "UserName="; UserName; "="
'        tt1 = InputBox("testing prompt", , , 4400, 4500)  'TESTING ONLY
        If strBuffer <> "STONEMAN" Then GoTo line_30    'vip todo comment out to activate
        If strBuffer = "STONEMAN" Then GoTo line_30     'february 01 2001
        If UserName = "stoneman" Then GoTo line_30
        If UCase(UserName) = "KENNETH" Then GoTo line_30    'february 02 2001
        If UCase(UserName) = "KEN M" Then GoTo line_30    'february 02 2001
        If strBuffer = "KEN" Then GoTo line_30          'february 02 2001
        MsgBox strBuffer + "=invalid computer " + UserName + "=username"
        
        GoTo End_32000                                  'february 01 2001
line_30:
        If date_check <> "YES" Then
            GoTo File_40    'no date check if disabled
        End If        'october 27 2000
        If ddemo = "YES" Then
            GoTo File_40    'no date check on demo version
        End If              'as it is restricted enough
'date check should be good till end of march 2001
        SSS = Format(Now, "ddddd ttttt")    'display today
        temp2 = InStr(SSS, "12")            '26Feb2012 if no 12 then not valid just for 2012
        If temp2 = 0 Then
        Print "invalid date="; SSS; "="
        tt1 = InputBox("software expired visit home site for new one", , , xx1 - offset1, yy1 - offset2) 'TESTING ONLY
            GoTo End_32000
        End If              '26Feb2012 no free lunch after 2012
'        temp2 = InStr(SSS, "/")
'        temp1 = InStr(temp2 + 1, SSS, "/2010") 'for 2010 only
'        Print "invalid date="; SSS; "="
'        tt1 = InputBox("year not /00", , , 4400, 4500)  'TESTING ONLY
'        If temp1 = 0 Then
'            temp1 = InStr(temp2 + 1, SSS, "/2000")
'        End If
'        If temp1 = 0 Then
'            temp1 = InStr(temp2 + 1, SSS, "/01")
'        End If
'        If temp1 = 0 Then
'            temp1 = InStr(temp2 + 1, SSS, "/2001")
'        End If
'to activate for windows 2000 need a proper date check here as with ME above
'        If temp1 = 0 Then
'        Print "invalid date="; SSS; "="
'        tt1 = InputBox("year not /10", , , xx1 - offset1, yy1 - offset2) 'TESTING ONLY
'            GoTo End_32000
'       End If              'year must be 2000
line_35:
        If Left(SSS, 1) = " " Then
           SSS = Right(SSS, Len(SSS) - 1)
            GoTo line_35
        End If
        If Left(SSS, 3) <> "10/" And Left(SSS, 3) <> "11/" _
            And Left(SSS, 3) <> "12/" And Left(SSS, 2) <> "1/" _
            And Left(SSS, 2) <> "2/" And Left(SSS, 2) <> "3/" Then
'             Then
        Print "software expired="; SSS; "="
        tt1 = InputBox("e-mail stonedan@telusplanet.net", , , xx1 - offset1, yy1 - offset2) 'TESTING ONLY
            GoTo End_32000
        End If              'must be march also
    screen_capture = "NO"           'march 15 2001

File_40:

    save_line = "40"
    Close #ExtFile          'march 18 2001
    copy_photo = "NO"       'march 18 2001

'09 december 2002    rand = 0                '18 august 2002
    text_pause = 0          '05 october 2002
    debug_photo = False     '12 october 2002
'    debug_photo = True     '01 february 2003 09 september 2003
If InStr(UCase(App.EXEName), "DEBUG") <> 0 Then
        debug_photo = True      '30Jun2012a
End If                                  '30Jun2012
    tt1 = ""
    encript = ""
    prev_option = ""
    If extract_yes = "YES" Then
        Close #ExtFile
        line_len = Val(Cmd(21))
    End If          'november 12 2000
    tot_s1 = 0      'january 01 2001
    tot_s2 = 0      'january 01 2001
    tot_s3 = 0      'january 01 2001
    tot_s4 = 0      '23 june 2002
    tot_s5 = 0      '23 june 2002
    tot_s6 = 0
    extract_yes = ""    'november 12 2000
    skip_info = ""      'november 14 2000
'    img_ctrl = "NO"           'february 21 2001
    img_ctrl = "YES"           'march 31 2001
    stretch_img = "YES"         'march 31 2001
'22 december 2002
'    If UCase(Cmd(38)) = "STRETCH" Then
'        stretch_info = " (STRETCH)" '21 december 2002
'    Else
'        stretch_info = " (NORMAL)"
'    End If
    dbltime1 = 0
    dbltime2 = 0
line_45:            'december 3 2000
    boundarycnt = 0     'december 31 2000
    boundarystr = ""    'december 31 2000
    mbxyes = ""         'december 17 2000
    dateskip = "F"      'december 23 2000
    GoSub InputFile_24000       'march 20/00
'        Print "tt1 SAVE_ttt="; tt1; "="; SAVE_ttt
'        ttt = InputBox("testing file_40", , , 4400, 4500)  'TESTING ONLY
    
    If tt1 = "" Then
        GoTo End_32000
    End If
    append_start1 = "append start"   'april 10 2001
    append_end1 = "append end"       'april 10 2001
    mbxi = 0            'december 17 2000
    emailsea = ""       'december 11 2000
    SAVE_ttt = "in"     'december 10 2000
    GoSub Control_28000 'december 10 2000
    extract_yes = ""        'december 31 2000
    photo_cnt = 0           'march 17 2001
    dirdates = ""           'april 11 2001
    hilite_hh = ""          'april 22 2001
    hilite_cnt = 0          'april 22 2001
    rand_str = ""           '18Mar2012
    If rand = -1 Then           '09 december 2002
        Randomize      'use system timer to start randomizer
        cnt = 99999999
        GoSub line_30920
'        rand_cnt = zzz_cnt - 3 'keep the total number of records in file
        rand_cnt = zzz_cnt - 6 'keep the total number of records in file 28 april 2003 (this fixed it) ver=1.04a
                                'seems to loop if it hits end of file
                                'Hint: (place 6 blank lines at end of file
                                'so last 3 pictures can be seen in random mode.
        zzz_cnt = 0
'        xtemp = InputBox(" testing doug  " + CStr(rand_cnt), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
        rand_no = Int(rand_cnt * Rnd + 1)   'this line moves to after hits or misses
        rand_str = "random#" + CStr(rand_no) '18Mar2012
    End If                      '09 december 2002
What_50:

    save_line = "50"    'for error handling
    pp_entered = ""     'november 6 2000
        SSS = Format(Now, "ddddd ttttt")    'display today
'deactivate the date check below have the program work for ever
   
    DoEvents
'    Err.Raise 6       'test this to give error info
                        ' works for all error numbers
 '   Beep        'test the beep bell

    prompt2 = ""
    ppoffset1 = 0       '27 july 2002
    ppoffset2 = 0       '27 july 2002
    'find out what can be called of use **vip** todo
    'test area for ggg = GetSetting("?","*")
'    ggg = GetSetting("RegCust", "Startup", "LastEntry", "0")
'        Print "ggg="; ggg; "="
'        tt1 = InputBox("testing What_50", , , 4400, 4500)  'TESTING ONLY
'    save_line = "50 testing"    'testing only
 '   ggg = LoadResData("C:|Program Files\Plus!\Themes\UNDERW~5.WAV", 6)

    
    
'        Print "ggg="; ggg; "="
'        tt1 = InputBox("testing What_50", , , 4400, 4500)  'TESTING ONLY

'*********************************************************
'main option entry here entry option prompt
'*********************************************************
'10 august 2002 allow for different keep options for photo file and text files
    If xxx_found = "YES" Then
        Cmd(20) = Cmd(41)
        Cmd(33) = "PHOTO"
'        tt1 = InputBox("testing 01 January 2005", , , 4400, 4500)  'TESTING ONLY
    End If
    If xxx_found = "NO" Then
        Cmd(20) = Cmd(42)
        Cmd(33) = "D"
    End If
    
'10 august 2002 use the cmd(20) and over-ride with above

'    If xxx_found = "NO" And UCase(Cmd(20)) <> "Q" Then
'        Cmd(20) = "C"
'        Cmd(33) = "D"
'    End If              'november 3 2000
'    If xxx_found = "YES" Then
'        Cmd(20) = "P"
'        Cmd(33) = "PHOTO"       'december 18 2000
'    End If              'november 3 2000
'10 august 2002 comment out the above logic

    iimport = ""        'december 24 2000
    If InStr(UCase(TheFile), "\MAIL\") <> 0 And SAVE_ttt = "in" Then Cmd(20) = "email" 'december 12 2000
    If InStr(UCase(TheFile), "\OUTLOOK EXPRESS\") <> 0 And SAVE_ttt = "in" Then Cmd(20) = "email" 'december 13 2000
    Set Picture = LoadPicture()     'lp#1
    page_prompt = ""        'december 31 2000
    MAX_CNT = Val(Cmd(1))        'december 31 2000
    Context_lines = Val(Cmd(22)) 'december 31 2000
    If Context_lines > 40 Then Context_lines = 40  'February 04 2001
    mess_cnt = 0                'january 03 2001
    array_pos = 0               'january 21 2001
    array_prt = 0
    temp_sec = -1               'march 14 2001
' WinSeek.Show           'february 26 2001
' WinSeek.SetFocus
    'february 26 2001
'    frmOption1.Caption = "Option prompt"    'february 26 2001
'    MDIOption1.Show      'february 26 2001
'    xtemp = InputBox("testing prompt_ after option1 show=" + TheFile + "*" + UCase(xtemp) + "*" + CStr(wrap_cnt) + " " + CStr(cnt), "test", , 4400, 4500) '
'    SetFocus           'february 26 2001
'    Unload frmOption1   'february 26 2001
    If screen_capture = "YES" Then
            new_delay_sec = 5           '03 September 2004
            GoSub line_30300            '03 September 2004
'        delay_sec = 5      'march 15 2001
'        GoSub line_30000
    End If
    dsp_cnt = 0            'may 12 2001
    dbltime2 = Timer      'get the end time may 12 2001
    Test1_str = ""
    If dbltime1 <> 0 And dbltime1 <> dbltime2 Then
        Test1_str = "  elap=" + Format(dbltime2 - dbltime1, "#####0.000")
    End If      'show elapsed times may 12 2001
    multi_prompt2 = ""  '04 September 2004
Get_no2:                'march 21 2002
    search_str = ""         '26 august 2002
'        xtemp = InputBox(" input prompt #2 multi_prompt2 " + multi_prompt2 + "*" + ttt, " testing Prompt #2   ", , xx1 - offset1, yy1 - offset2)
    If Cmd(20) = "Y" And SAVE_ttt <> "in" Then Cmd(20) = "" '15 december 2004
    If Len(multi_prompt2) > 0 Then
        temp1 = InStr(1, multi_prompt2, sep)
        If temp1 <> 0 Then
            ttt = Left(multi_prompt2, temp1 - 1)
            multi_prompt2 = Right(multi_prompt2, Len(multi_prompt2) - temp1)
'            xtemp = InputBox(" input prompt #2* multi_prompt2 " + multi_prompt2 + "*" + ttt, " testing Prompt #2*   ", , xx1 - offset1, yy1 - offset2)
        Else
            ttt = multi_prompt2
            multi_prompt2 = ""
'            xtemp = InputBox(" input prompt #2** multi_prompt2 =" + multi_prompt2 + "*" + ttt, " testing Prompt #2**   ", , xx1 - offset1, yy1 - offset2)
 '           Cmd(20) = ""    '05 September 2004
        End If
        GoTo auto_p2        '04 September 2004
    
    End If                  '04 September 2004      allow for multiple inputs at a time

'        xtemp = InputBox(" input prompt #2 test " + auto_exe + "*", " testing Prompt #2   ", , xx1 - offset1, yy1 - offset2)
'12 September 2004    If UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\" Then
'    If (UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.Path + App.EXEName), "BACKGRD") <> 0 Then
'19Aug2011    If (UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.EXEName), "BACKGRD") <> 0 Then
        mixx = False        '12Feb2017
If InStr(UCase(App.EXEName), "MIX") <> 0 Then
        mixx = True
End If                                  '12Feb2017
        budgeyn = False        '17Dec2017
If InStr(UCase(App.EXEName), "BUDGE") <> 0 Then
        budgeyn = True
End If                                  '17Dec2017
If Left(UCase(App.EXEName), 6) = "RANDOM" Then
        rand_prog = True
        rand = True
        If mixx Then Cmd(81) = "RAND_GROUP   " '29Oct2012 if random and mix then rand_group needed
End If          '29Oct2012
If InStr(UCase(App.EXEName), "ALT") <> 0 Then
        altt = True
End If                                  '12Aug2012
'If Left(UCase(App.EXEName), 7) = "BIGTEXT" Then
'            xtemp = InputBox("testing#4dd Match 29Oct2012 rand=" + CStr(rand) + " break_num=" + CStr(break_num) + " group_match=" + CStr(group_match) + " rand_no=" + CStr(rand_no) + " zzz_cnt=" + CStr(zzz_cnt) + " rand_prog=" + CStr(rand_prog) + " aaa=" + Left(aaa, 10), , , 4400, 4500) 'TESTING ONLY
If Left(UCase(App.EXEName), 7) = "BIGTEXT" Or mixx = True Then      '30Jun2012
        ttt = "C"
        If mixx = True Then
            ttt = "WW"       '30Jun2012b this seems to be the problem area
        End If
'            tt1 = InputBox("testing#2a 29Oct2012 mixx=" + CStr(mixx) + " ttt=" + ttt + " rand_prog=" + CStr(rand_prog), , , 4400, 4500) 'TESTING ONLY
        If Left(UCase(Cmd(77)), 5) <> "CHARA" Then      '**note** do not override what they have if not noCHARA
'        xtemp = InputBox(" 29Oct2012 input prompt #2 mixx= " + CStr(mixx) + "*", " testing Prompt #2   ", , xx1 - offset1, yy1 - offset2)
            MAX_CNT = 10       'lines per screen cmd(1)
            Font.Size = 48  'font size cmd(2) 48 point size
            BackColor = 0   'black background cmd(3)
            Def_Fore = 15   'white text cmd(4)
            Cmd(5) = 14     'hi-lite color yellow
'            ForeColor = QBColor(Def_Fore)
            'Cmd(21) = 26
            line_len = 260   'cmd(21)  make line length long.. all formatting done by user
            Cmd(31) = ":"   'the highlite this element
            hilite_this = ":" 'cmd(31)
            Context_lines = 0  'cmd(22)
            rand = 0            '24Jun2012
            Cmd(76) = "LINEPAUSE==0.15"
            Cmd(77) = "CHARACTERPAUSE==.0333"
            Cmd(80) = "PROMPTDETAILS"
                    '22Jun2012
            'frmproj2.AutoRedraw = True
        End If                                      'ie they have their own settings no need to override
        GoTo auto_p2
End If                  '21Jun2012
If Left(UCase(App.EXEName), 14) = "CATMYDRIVEGOLF" Then
        ttt = "GOLF"
        GoTo auto_p2
End If                  '18Jun2012
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        ttt = "GF"
        GoTo auto_p2
End If                  '18Jun2012
    If Left(UCase(App.EXEName), 6) = "RANDOM" Or Left(UCase(App.EXEName), 10) = "SEQUENTIAL" Then
        ttt = "WW"
        Cmd(49) = "NORANDOM"
        If Left(UCase(App.EXEName), 6) = "RANDOM" Then
            Cmd(49) = "RANDOM"
            If UCase(Left(Cmd(81), 10)) = "RAND_GROUP" Then
                frmproj2.Caption = "turn cmd(81) RAND_GROUP off if no BREAK records" '12May2012
            End If
        End If
        GoTo auto_p2                '12May2012
    End If
    If (InStr(1, UCase(Cmd(45)), UCase(App.EXEName)) <> 0 And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.EXEName), "BACKGRD") <> 0 Then
        ttt = RTrim(Cmd(47))       'the search options entered here for now use "WW"
        'want to make it so multiple prompts can be done todo **vip**
'            xtemp = InputBox(" backgrd test2", " testing Prompt #2*   ", , xx1 - offset1, yy1 - offset2)
        If debug_photo Then     '12 october 2002
            xtemp = InputBox("DOUG TESTING AUTO PROMPT #2" + ttt, , , 4400, 4500) 'TESTING ONLY
        End If
        GoTo auto_p2
    End If                  '07 december 2002
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then        '27 October 2004
            ttt = Right(Cmd(71), Len(Cmd(71)) - 11)
'            xtemp = InputBox("27 october 2004 TESTING 1 =" + ttt, , , 4400, 4500) 'TESTING ONLY
            BatchFile = FreeFile
            Open ttt For Input As #BatchFile
            DoEvents
            
        ttt = "RRR"                 'testing only re the close of the file etc...
        Line Input #BatchFile, ttt
'            xtemp = InputBox("27 october 2004 TESTING 1a =" + ttt, , , 4400, 4500) 'TESTING ONLY
'    interrupt_prompt2 = UCase(ttt) '20 November 2004
        GoTo auto_p2
    End If                                                  '27 October 2004
'19 august 2003 the following line deactivated. If xxx. not found then disable random. We want it random here
'19 August 2003    If xxx_found = "NO" Then rand = 0 '19 january 2003
                        'the above can be over-ridden below
   
   ttt = InputBox("P P1 P2 TT WW SS CH X RAND NORAND RANDA NORANDA THUMB HELP option ", "Option Prompt #2  " + TheFile + Test1_str, UCase(Cmd(20)), xx1 - offset1, yy1 - offset2)

'    frmproj2.BorderStyle = "0"            '25 november 2002
'    frmproj2.BorderStyle = "0"            '25 november 2002
'10 august 2002 allow for cmd(41) and cmd(42) to override
'    interrupt_prompt2 = UCase(ttt) '20 November 2004
auto_p2:                    '07 december 2002
    ttt = UCase(ttt)        '07 december 2002
    
    If inin = "" Then
        start_point = 0  '21 February 2007
        start_point = 10 '11 March 2007
        line_start_point = start_point  '21 February 2007
        keep_line_start_point = line_start_point    '25Dec2011
        cont_str = ""           '08Feb2012
    End If              '21 February 2007
        cont_str = "S"      '16Apr2012
'04 September 2004 below
    If InStr(1, ttt, sep) <> 0 Then
        temp1 = InStr(1, ttt, sep)
        multi_prompt2 = Right(ttt, Len(ttt) - temp1)
        ttt = Left(ttt, temp1 - 1)
'        xtemp = InputBox(" input prompt #2a multi_prompt2 " + multi_prompt2 + "*" + ttt, " testing Prompt #2a   ", , xx1 - offset1, yy1 - offset2)
    End If                  '04 September 2004
    If Left(ttt, 6) = "SPEED=" Then
        play_speed = Val(Right(ttt, Len(ttt) - 6))
        keep_play_speed = play_speed        '08Nov2011
        GoTo Get_no2
    End If                      '29 October 2003 testing this a bit
    
    p2p2 = UCase(ttt)       '10 august 2002
    If ttt = "SS" And ss_only <> "YES" Then
        ss_only = "YES"
        GoTo Do_Search_110
    End If                  '07 december 2002
    If p2p2 = "DEBPHO" Then
        debug_photo = True
        GoTo Get_no2
    End If                  '12 october 2002 allow for debug of photo problems
    
'16 July 2003 testing only below  28Dec2011 this is for batch start point other is per video and audio file
    If Left(p2p2, 7) = "START==" Then
        start_point = Val(Right(p2p2, Len(p2p2) - 7))
        line_start_point = start_point  '22 March 2004
        keep_line_start_point = line_start_point    '28Dec2011
        GoTo Get_no2
    End If                  '16 July 2003 Ver=1.07T testing the start point
    
    If Left(p2p2, 4) = "ELAP" Then
        elapse_yn = "YES"         '13 July 2003
        GoTo Get_no2
    End If                      '13 July 2003
    If p2p2 = "PAUSE" Or p2p2 = "PA" Then
        text_pause = True
        extract_yes = "NO"      'do not want these two taking up disc space
        auto_redraw = "NO"
        frmproj2.AutoRedraw = False
        inin = ""
        GoTo Get_no2
    End If                  '05 october 2002
'12May2012
'12May2012    If p2p2 = "RAND" Then
    If p2p2 = "RAND" Or Left(UCase(App.EXEName), 6) = "RANDOM" Then
        rand = -1
        Randomize      'use system timer to start randomizer
        cnt = 99999999
        GoSub line_30920
        rand_cnt = zzz_cnt - 3 'keep the total number of records in file
                                'seems to loop if it hits end of file
        zzz_cnt = 0
'        xtemp = InputBox(" testing doug  " + CStr(rand_cnt), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
        rand_no = Int(rand_cnt * Rnd + 1)   'this line moves to after hits or misses
        rand_str = "random#" + CStr(rand_no) '18Mar2012
'        If rand Then
'            xtemp = InputBox(" testing doug randomizer rnd  " + CStr(rand_no), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
'        End If
            GoSub Control_28000
'22 december 2002
            Cmd(49) = "RANDOM"          '09 December 2002
'            random_info = " (RANDOM)"     '21 december 2002
            GoSub line_30800    'kill and update the control file
            frmproj2.Caption = program_info + random_info + stretch_info '09 december 2002
            If Left(UCase(App.EXEName), 6) <> "RANDOM" Then GoTo Get_no2     '12May2012
            p2p2 = Cmd(20)         '12May2012
            ttt = Cmd(20)          '12May2012
'12May2012        GoTo Get_no2
'12May2012 test below
'        If rand Then
'            xtemp = InputBox(" testing doug randomizer rnd  " + CStr(rand_no), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
'        End If
    End If                  '18 august 2002
            If Left(UCase(App.EXEName), 9) = "RANDOMSEG" Then
                thumb_nail = "YES"
                rand1 = -1
                Randomize
                Cmd(65) = "RANDBEG"
            End If      '23Jun2012a
'12May2012    If p2p2 = "NORAND" Then
    If p2p2 = "NORAND" Or Left(UCase(App.EXEName), 10) = "SEQUENTIAL" Then
            rand = 0
            GoSub Control_28000
            Cmd(49) = "NORANDOM"    '09 december 2002
'22 december 2002
'            random_info = ""        '21 december 2002
            GoSub line_30800    'kill and update the control file
            frmproj2.Caption = program_info + stretch_info '09 december 2002
        If Left(UCase(App.EXEName), 10) <> "SEQUENTIAL" Then GoTo Get_no2       '12May2012
            p2p2 = Cmd(20)     '12May2012
            ttt = Cmd(20)     '12May2012
'12May2012        GoTo Get_no2
    End If              '09 december 2002
    
    If p2p2 = "DETAIL" Then
            GoSub Control_28000
            Cmd(56) = "PHOTO_DETAIL"    '18 November 2004
            GoSub line_30800    'kill and update the control file
            detailyn = "PHOTO_DETAIL"    '18 November 2004
        GoTo Get_no2
    End If              '18 November 2004
    
    If p2p2 = "NODETAIL" Then
            GoSub Control_28000
            Cmd(56) = "noPHOTO_DETAIL"    '18 November 2004
            GoSub line_30800    'kill and update the control file
            detailyn = "noPHOTO_DETAIL"    '18 November 2004
        GoTo Get_no2
    End If              '18 November 2004
    
    If p2p2 = "RG" Then
            GoSub Control_28000
            Cmd(81) = "RAND_GROUP"    '09Feb2012
            GoSub line_30800    'kill and update the control file
        GoTo Get_no2
    End If              '09Feb2012
    
    If p2p2 = "NORG" Then
            GoSub Control_28000
            Cmd(81) = "noRAND_GROUP"    '09Feb2012
            GoSub line_30800    'kill and update the control file
        GoTo Get_no2
    End If              '09Feb2012
    

    If Left(p2p2, 4) = "THUM" Then
        thumb_nail = "YES"          '16 July 2003 ver=1.07T
        GoSub Control_28000
        Cmd(67) = "THUMB"           '14 April 2004
        GoSub line_30800
        GoTo Get_no2
    End If

    If Left(p2p2, 6) = "NOTHUM" Then
        thumb_nail = "NO"
        GoSub Control_28000
        Cmd(67) = "noTHUMB"           '14 April 2004
        GoSub line_30800
        GoTo Get_no2
    End If

'23 March 2004
    If p2p2 = "RANDA" Then
        rand1 = -1
        Randomize
        GoSub Control_28000
        Cmd(65) = "RANDBEG"
        GoSub line_30800
        GoTo Get_no2
    End If
    
    If p2p2 = "NORANDA" Then
        rand1 = 0
        GoSub Control_28000
        Cmd(65) = "noRANDBEG"
        GoSub line_30800
        GoTo Get_no2
    End If                  '23 March 2004 add random start point in video
    
    
    If p2p2 = "NOVIDEO" Then
            GoSub Control_28000
            Cmd(53) = "NOSHOWVIDEO"    '10 FEBRUARY 2003
            GoSub line_30800    'kill and update the control file
            videoyn = "NOSHOWVIDEO"
        GoTo Get_no2
    End If              '10 FEBRUARY 2003
    
    If p2p2 = "VIDEO" Then
            GoSub Control_28000
            Cmd(53) = "SHOWVIDEO"    '10 FEBRUARY 2003
            GoSub line_30800    'kill and update the control file
            videoyn = "SHOWVIDEO"
        GoTo Get_no2
    End If              '10 FEBRUARY 2003
    
    '05 december 2002 If p2p2 = "P1" Or p2p2 = "P2" Or p2p2 = "WW" Or p2p2 = "SS" Then
 'midway thru the program appx
     
    If p2p2 = "P" Or p2p2 = "WW" Or p2p2 = "SS" Then
        If Cmd(41) <> p2p2 Then
            GoSub Control_28000
            Cmd(41) = p2p2
            GoSub line_30800    'kill and update the control file
        End If
    End If                  '10 august 2002
    
'15 December 2004     If p2p2 = "C" Or p2p2 = "S" Or p2p2 = "Q" Or p2p2 = "CC" Then
    If p2p2 = "C" Or p2p2 = "S" Or p2p2 = "Q" Or p2p2 = "CC" Or p2p2 = "Y" Then
        If Cmd(42) <> p2p2 Then
            GoSub Control_28000
            Cmd(42) = p2p2
            GoSub line_30800    'kill and update the control file
        End If
    End If                  '10 august 2002
    If p2p2 = "CC" Then
        search_str = "CC"   '26 august 2002
        p2p2 = "S"
        prompt2 = "S"
        ttt = "S"
    End If
    
    dbltime2 = Timer  'may 12 2001
    dbltime1 = Timer    'Get the start time may 12 2001
    prev_ttt = ttt      'april 10 2001
    If UCase(ttt) = "Z" Then prev_option = ""   'february 25 2001
    If prev_option = "CH" Then
        SetFocus            'this solved the problem of the esc / enter not interrupting
    End If
                            'the program once the "CH" option was done february 20 2001
line_60:            'february 09 2001
    tt1 = ttt           'november 14 2000
    ttt = UCase(ttt)    'all to upper case june 14/99
    If ttt = "M" Then
        frmproj2.WindowState = 1    '0=normal 1 = min 2 = max
        'when I used co-ordinates of 40000 , 40000 below the input box failed and
        'it skipped immediately to the get_no2 prompt what the hey
        ttt = InputBox(" window state prompt", "minimized  ", "WHAT", 20000, 20000)
        frmproj2.WindowState = 2    '
        GoTo Get_no2
    End If              'march 21 2002
' "Y" for yesterday display in text file december 09 2001
' read till end of file then do similar to the "B" option
        If ttt = "Y" Then
            save_line = "65"
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            DoEvents       'yield to operating system
            dblStart = Timer    'get the start time

line_65:
            Line Input #OutFile, aaa
            zzz_cnt = zzz_cnt + 1
            GoTo line_65
        End If

    If ttt = "HELP" Then
        rand = 0        '10 december 2002
        Close #OutFile
        DoEvents
        DoEvents
        GoSub line_30900
        Close #OutFile
        DoEvents
        OutFile = FreeFile
        Open "help.txt" For Input As #OutFile
        DoEvents
        prompt2 = "C"
        printed_cnt = 1 'otherwise the "reading 1" shows up on screen
        inin = "A"    'forces all data to be displayed
        SSS1 = ""       'when set to "A" only lines with "A" in displayed
        SAVE_KEEPS1 = "=="        'the search match is an and (just use one at a time)
        SAVE_KEEPS2 = "CMD("                '**vip** the search here must be in all capitals
        SAVE_KEEPS3 = ") "            'use this for some other display all caps
        SAVE_KEEPS4 = ""        '09 JUNE 2002
        SAVE_KEEPS5 = ""
        SAVE_KEEPS6 = ""
        SAVE_ttt = "C"      'this ensures it is "C" context display
        Picture_Search = ""     'this one was important in making the help work
        hi_lites = "YES"
'    frmproj2.Caption = program_info + " (cls #1)" '18Dec2013
        
        Cls             'clear the screen
' can not seem to get this routine to display the hilites on the second pass
' after a photo has been displayed and the option prompt "help" request
' not a big deal but just having no luck today october 22 2001
'        SSS = SAVE_KEEPS1 + sep + SAVE_KEEPS2 + sep + SAVE_KEEPS3
'       xtemp = InputBox("DOUG 1 TESTING SSS1 SSS2 SSS3" + SSS1 + SSS2 + SSS3, , , 4400, 4500)         'TESTING ONLY
'        xtemp = InputBox("DOUG 2 TESTING SAVE_ttt hi_lites" + SAVE_ttt + "," + hi_lites, , , 4400, 4500)  'TESTING ONLY
'        KEEPS1 = ""
'        KEEPS1 = ""
'        KEEPS2 = ""
'        KEEPS3 = ""
'        xtemp = InputBox("DOUG 3 TESTING KEEPS1 KEEPS2 KEEPS3 " + KEEPS1 + "," + KEEPS2 + "," + KEEPS3, , , 4400, 4500) 'TESTING ONLY
'        xtemp = InputBox("DOUG 4 TESTING tt1" + tt1, , , 4400, 4500) 'TESTING ONLY
'        tt1 = ""
'        xtemp = InputBox("DOUG 5 TESTING save_search" + save_search, , , 4400, 4500) 'TESTING ONLY
'        xtemp = InputBox("DOUG 6 TESTING SAVE_KEEPS1" + SAVE_KEEPS1, , , 4400, 4500) 'TESTING ONLY
'        xtemp = InputBox("DOUG 7 TESTING SAVE_KEEPS2" + SAVE_KEEPS2, , , 4400, 4500) 'TESTING ONLY
'        line_match = "Y"
        GoTo input_1000a
        
    End If              'october 14 2001
    If ttt = "DIRDATES" Then
        dirdates = "Y"
        GoTo What_50
    End If              'april 11 2001
    If Left(ttt, 2) = "HH" Then
        hilite_this = Mid(prev_ttt, 3)
        GoSub Control_28000         '19 january 2003
        Cmd(31) = Mid(prev_ttt, 3)  ' only hilites data not on matching line
        GoSub line_30800            '19 january 2003
        hilite_hh = "Y"             'april 22 2001
        GoTo What_50
    End If              'april 10 2001
    If Left(ttt, 3) = "HLL" Then
        append_end1 = Mid(prev_ttt, 4)
        GoTo What_50
    End If              'april 10 2001
    If Left(ttt, 2) = "HL" And Mid(ttt, 3, 1) <> "L" Then
        append_start1 = Mid(prev_ttt, 3)
        GoTo What_50
    End If              'april 10 2001
    If ttt = "P1" Then
        img_ctrl = "YES"       'march 31 2001
        stretch_img = "YES"     'march 31 2001
'22 december 2002
'        stretch_info = " (STRETCH)" '21 december 2002
'        frmproj2.Caption = program_info + random_info + stretch_info '09 december 2002
        GoSub Control_28000        'october 07 2001 read the most current then update
        Cmd(38) = "stretch"
        search_prompt = "in"        'october 07 2001
        GoSub line_30800        'september 23 2001
        frmproj2.Caption = program_info + random_info + stretch_info '22 december 2002
        If debug_photo Then     '12 october 2002
            xtemp = InputBox("DOUG TESTING photo 1" + ttt, , , 4400, 4500) 'TESTING ONLY
        End If
        ttt = "P1"              'OCTOBER 07 2001
        ttt = "P"              '05 december 2002
        GoTo What_50            '05 december 2002
    End If                  'march 31 2001
    If ttt = "P" Then
        ttt = "P1"
    End If                  'september 23 2001
        
    If ttt = "SC" Then
        screen_capture = "YES"  'march 15 2001
        GoTo What_50
    End If
    If ttt = "DIR" Then
        ttt = "GF"              'february 18 2002
        diryes = "DIR"
        dirdates = "Y"
    End If                      'february 18 2002
    Mergem = "NO"               'february 24 2002
    If ttt = "MERGE" Then
        ttt = "GF"              'february 24 2002
        diryes = "DIR"
        dirdates = "Y"
        Mergem = "YES"
    End If                      'february 24 2002
    
    If ttt = "GOLF" Then
            ttt = "GF"          '25Jun2010
            Mergem = "GOLF"     '25Jun2010
    End If
     If ttt = "VL" Then
            ttt = "GF"          '04Feb2012   put the video_length and size in bytes on the catalog
            Mergem = "VL"       '04Feb2012
    End If
    rewind = "NO"   '12Apr2014
    If Left(ttt, 2) = "BB" Then
        ttt = "FF" + Right(ttt, Len(ttt) - 2) '12Apr2014 rewind very similar to fast forward
        rewind = "YES"
'        xtemp = InputBox("testing 12Apr2014 " + ttt, , , 4400, 4500) '12Apr2014
    End If      '12Apr2014
    If Left(ttt, 2) = "FF" Then 'ie FF.7 for a .7 sec play   17Mar2014
        nocount = 0 '12Apr2014
        fast_forward = 0.7         '17Mar2014 if none entered play for .7 of a second
        If Len(ttt) > 2 Then fast_forward = Val(Right(ttt, Len(ttt) - 2)) 'set a new delay time on screen saver
        ttt = "GF"          '06Feb2012  catalog for fast forward of videos
        Mergem = "FF"       '06Feb2012
    End If
    
    If ttt = "GF" Then
            GoSub Control_28000
            rand = 0            '19 July 2003 take random off GF option
            Cmd(49) = "NORANDOM"
            GoSub line_30800    'kill and update the control file
                                '19 July 2003 (somehow having random set causes the tag record
                                'on mp3 files not to be read properly (so this should be a temp fix)
            GoSub Control_28000  '19 July 2003
        Close #OutFile          '18 August 2004 so there is no conflict with file opened at prompt number 1 one
        GoSub line_30700            'april 01 2001
'15 december 2002        GoTo What_50
        GoTo File_40
    End If
    
    If Left(ttt, 2) = "PC" Then
        photo_cnt = Val(Mid(ttt, 3)) - 1
        GoTo What_50
    End If              'march 17 2001
    If ttt = "CP" Then
        copy_photo = "YES"      'march 17 2001
        'need to allow for change of file name etc here
        xtemp = Cmd(19)
        GoSub line_16000        '
        If UCase(FileExt) = UCase(TheFile) Then
            ttt = "X"
            GoTo File_40
        End If
        'might want to make the open below append or new
        
'            save_line = "16100-1"  '18 March 2007 testing only only
            GoSub line_16100    'open the replace.txt for output
        xtemp = InputBox(" output directory (you can rename folder later)", "Directory Prompt   ", Cmd(23), xx1 - offset1, yy1 - offset2)
        temps = xtemp
        good_fold = xtemp   '15Aug2016
'        good_fold = ""      '28Aug2016 they seem to be going to the default on c: somehow
'        If Len(temps) < 4 Then
'          GoTo What_50
'        End If                  'april 08 2001
'09Aug2016 a blank directory is allowed because those with mp3 errors blow will be ougput
        badmp3_fold = InputBox(" problems directory (you can rename folder later) " + Cmd(84), "Directory Prompt " + Test1_str + " " + zemp, "", xx1 - offset1, yy1 - offset2) '08Aug2016
'        badmp3_fold = "d:\bad_mp3\"        'JUST DURING TESTING 09Aug2016
'        GoTo What_50        '08Aug2016 testing only just to check test1_str above
save_dir:
        photo_dir = xtemp 'ie c:\search\tempfold\
        xtemp = InputBox(" keep original file name Y/N  <Y> (pict for pict00001 & pict00002 etc)", "Output file name prompt   ", "Y", xx1 - offset1, yy1 - offset2)
        photo_file = xtemp 'ie pict
        If xtemp = "Y" Then photo_file = ""
'        xtemp = InputBox("testing dir" + photo_dir, , , 4400, 4500) 'TESTING ONLY
'        xtemp = InputBox("testing file" + photo_file, , , 4400, 4500) 'TESTING ONLY
        GoTo What_50
    End If          'end "CP" if statement
    prev_option = ttt   'february 25 2001
    If ttt = "P2" Then
        stretch_img = "NO"      'march 31 2001
'22 december 2002
'        stretch_info = " (NORMAL)"       '21 december 2002
'        frmproj2.Caption = program_info + random_info + stretch_info     '25 november 2002
'        img_ctrl = "YES"        'february 21 2001
        img_ctrl = "NO"        'march 31 2001
        ttt = "P1"
        GoSub Control_28000       'october 07 2001
        Cmd(38) = "normal"      'september 23 2001
        search_prompt = "in"        'october 07 2001
        GoSub line_30800        'september 23 2001
        frmproj2.Caption = program_info + random_info + stretch_info     '25 november 2002
'        xtemp = InputBox("TESTING DOUG photo 2" + ttt, , , 4400, 4500) 'TESTING ONLY
        ttt = "P1"              'OCTOBER 07 2001
        ttt = "P"              '05 december 2002
        GoTo What_50            '05 december 2002
    End If
    If Cmd(20) = "email" And ttt = "IMPORT" Then
        emailsea = "Y"
        iimport = "Y"       'december 24 2000
        ttt = "F"           'for flash display
        extract_yes = "YES"
        xtemp = Cmd(19)
        GoSub line_16000        'get file name
        If UCase(FileExt) = UCase(TheFile) Then
            ttt = "X"
            GoTo File_40
        End If              'january 22 2001
'            save_line = "16100-2"  '18 March 2007 testing only only
        GoSub line_16100        'january 01 2001
    End If
    Cmd(20) = "" 'december 7 2000
    xxx_found = "" 'december 7 2000
    If ddemo = "YES" And ttt = "" Then
        GoTo line_80
    End If

'demo only allows the P1 and the SS options
    If ddemo = "YES" And ttt <> "P1" And ttt <> "X" And ttt <> "RRR" And ttt <> "SS" Then
        GoTo What_50
    End If
'december 11 2000 re the emailsea option email
    If ttt = "EMAIL" Then
        emailsea = "Y"
        ttt = "C"
    End If
    If ttt = "R" Then
        page_prompt = "NO"
        extract_yes = "YES"
        xtemp = Cmd(19)
        GoSub line_16000
        If UCase(FileExt) = UCase(TheFile) Then
            ttt = "X"
            GoTo File_40
        End If              'january 22 2001
'            save_line = "16100-3"  '18 March 2007 testing only only
        GoSub line_16100    'january 01 2001
        ttt = "Q"
    End If          'december 31 2000
    If ttt = "CROP" Then
        GoSub line_30600        'crop the input file to a given length
        SAVE_ttt = ""
        GoTo File_40
    End If          'december 31 2000
    If ttt = "DISC" Then
        MsgBox strMessage
        strMessage = App.Path + "*" + App.EXEName + " serial#= " + ser_num '07 december 2002
        MsgBox strMessage
        GoTo What_50
    End If                          'january 31 2001
    If ttt = "D" Then
        page_prompt = "NO"
        extract_yes = "YES"
        xtemp = Cmd(19)
        GoSub line_16000
'        xtemp = InputBox("testing prompt_=" + TheFile + "*" + UCase(xtemp) + "*" + CStr(wrap_cnt) + " " + CStr(cnt), "test", , 4400, 4500) '
        If UCase(TheFile) = UCase(FileExt) Then
            ttt = "X"
            GoTo File_40
        End If              'january 22 2001
'            save_line = "16100-4"  '18 March 2007 testing only only
        GoSub line_16100    'january 01 2001
'january 01 2001        MAX_CNT = 5
'january 01 2001        Context_lines = 2
        MAX_CNT = context_win
        Context_lines = Int((context_win - 1) / 2)
        If Context_lines > 40 Then Context_lines = 40   'january 28 2001
        ttt = "C"
    End If          'december 31 2000
   
    If ttt = "T" Then
        page_prompt = "NO"
        extract_yes = "YES"
        xtemp = Cmd(19)
        GoSub line_16000
        If UCase(FileExt) = UCase(TheFile) Then
            ttt = "X"
            GoTo File_40
        End If              'january 22 2001
'            save_line = "16100-5"  '18 March 2007 testing only only
        GoSub line_16100    'january 01 2001
        ttt = "S"   'see logic below for changing S to Q for search
    End If          'january 01 2001
'november 14 2000 the following stuff moved up to do_search area
    uppercase = "N"     'december 8 2000
    If ttt = "S" Then
        uppercase = "Y"
        ttt = "Q"
    End If              'december 8 2000
    If ttt = "XXX" Then
        extract_yes = "YES"     'november 12 2000
        text_pause = 0          '05 october 2002 don't want these two taking up disc
        xtemp = Cmd(19)
        GoSub line_16000        'get the file name november 17 2000
        If UCase(FileExt) = UCase(TheFile) Then
            ttt = "X"
            GoTo File_40
        End If              'january 22 2001
'            save_line = "16100-6"  '18 March 2007 testing only only
        GoSub line_16100    'january 01 2001

            GoTo What_50
    End If
    If ttt = "LL" Then
        line_len = Val(Cmd(21))
        GoTo What_50
    End If                  'december 6 2000
    If Left(ttt, 2) = "LL" Then
        line_len = Val(Right(ttt, Len(ttt) - 2))
        GoTo What_50
    End If                  'december 6 2000
    If Left(ttt, 7) = "SHOWPOS" Then
        showpos = "Y"
        GoTo What_50
    End If                  'december 6 2000
    '**keep** the showpos is handy to check the wrap functions for problems
    'and there will be problems
    If Left(ttt, 7) = "SHOWASC" Then
        showasc = "Y"
        GoTo What_50
    End If                  'december 6 2000
    If Left(ttt, 4) = "SKIP" Then
        skip_info = Right(tt1, Len(tt1) - 4)
        GoTo What_50                        'november 14 2000
    End If
    
    If Left(ttt, 2) = "NS" And Len(ttt) > 2 Then
        Cmd(28) = " " + Cmd(28) + " " + Right(ttt, Len(ttt) - 2) + " "
        GoSub line_29100  'set up new noshow elements
        GoTo What_50
    End If
'november 03 2000 allow switch between two control files ie for font sizes etc
    If Left(ttt, 3) = "CCC" Then
'        control_file = "c:\control1.txt"
        control_file = "control1.txt"   'december 3 2000 hint use control1.txt to switch multiple settings at once
        GoSub Control_28000
        DoEvents
'        control_file = "c:\control.txt"
'23 september 2002        control_file = "control.txt"       'december 3 2000
        SAVE_ttt = "in"     'november 23 2000
        GoTo What_50
    End If
    If Left(ttt, 2) = "SS" And Len(ttt) > 2 Then
        Cmd(26) = " " + Right(ttt, Len(ttt) - 2) + " "
'        ss_search = Right(ttt, Len(tt) - 2) + " " 'october 24 2000
        GoSub line_29200  'set up new screen saver elements
'        sscreen_saver = "Y" 'october 24 2000
'        prompt2 = "SS"
 '       Print "screensave(?)="; "*"; screensave(screencount); "*"; aaa; screencount; sscreen_saver; prompt2
 '       tt1 = InputBox("testing screen saver logic", , , 4400, 4500)  'TESTING ONLY
            
        ttt = "SS"
    End If

'28 april 2002 WW screen saver option added
    If Left(ttt, 2) = "WW" Then
        sscreen_saver_ww = "YES"
        interrupt_prompt2 = ""      '23 November 2004
        inin = ""
        ttt = "SS"
        Picture_Search = "YES"      '26 November 2004 testing this doug
'        ShowCursor = False        '29 november 2002
'        Me.MousePointer = False     '29 november 2002
    End If
    
    If Left(ttt, 6) = "SMLBUD" Then      '17Dec2017
        budgeyn = True
        If Len(ttt) > 6 Then       '17Dec2017
        GoSub Control_28000
        smlbud = Val(Right(ttt, Len(ttt) - 6))
        Cmd(85) = Format(smlbud, "####0.0000")
        If smlbud <> 0 Then GoSub line_30800     'do not update file if zero but let em try it
        End If
        GoTo Get_no2
    End If                       '17Dec2017
    If Left(ttt, 6) = "BIGBUD" Then      '17Dec2017
        budgeyn = True
        If Len(ttt) > 6 Then       '17Dec2017
        GoSub Control_28000
        bigbud = Val(Right(ttt, Len(ttt) - 6))
        Cmd(86) = Format(bigbud, "####0.0000")
'        If bigbud <> 0 Then GoSub line_30800     'do not update file if zero but let em try it
        End If
        GoTo Get_no2
    End If                       '17Dec2017
     
    If Left(ttt, 6) = "MAXBUD" Then      '17Dec2017
        maxbud = 4                     'when budge is in effect this must be non zero
        budgeyn = True
        If Len(ttt) > 6 Then       '17Dec2017
        GoSub Control_28000
        maxbud = Val(Right(ttt, Len(ttt) - 6))
        Cmd(87) = Format(maxbud, "####0.0000")
        If maxbud <> 0 Then GoSub line_30800     'do not update file if zero but let em try it
        End If
        GoTo Get_no2
    End If                       '17Dec2017
        
'         tt1 = InputBox("testing1 delay " + ttt + " " + Cmd(85), , , 4400, 4500) 'TESTING 22 March 2004


     If Left(ttt, 2) = "TT" And Len(ttt) > 2 Then
        GoSub Control_28000
        delay_sec = Val(Right(ttt, Len(ttt) - 2)) 'set a new delay time on screen saver
'february 08 2002 the 4 lines below were added to save the delay time
        hold_sec = delay_sec
        Cmd(27) = Format(hold_sec, "###0.0000")
        delay_sec = hold_sec            '29 november 2002
        GoSub line_30800
        delay_sec = hold_sec            '18 November 2004
 '        tt1 = InputBox("testing delay " + ttt + " " + Cmd(27), , , 4400, 4500) 'TESTING 22 March 2004
        GoTo Get_no2        '05 September 2004
'05 September 2004        GoTo What_50
    End If
    
     If Left(ttt, 6) = "OFFSET" And Len(ttt) > 6 Then
        GoSub Control_28000
        OffSet = Val(Right(ttt, Len(ttt) - 6)) 'set a new OffSet value
        Cmd(83) = "OFFSET==" + Format(OffSet, "###0.0000")
        GoSub line_30800
        GoTo Get_no2        '02Nov2011
    End If
     
    
    
    If ttt = "" And SAVE_ttt = "in" Then
        ttt = UCase(Cmd(20))                  'april 14/00
        If ttt = " " Then
            ttt = "C"
            Cmd(20) = ttt
        End If
    End If                  'default to "c" on incomming
    prompt2 = ttt + ""
    

    previous_count = 0
    If ttt = "SS" Then
        ttt = "P1"
'14May2012        If strecth_img = "NO" Then img_ctrl = "NO"    'march 31 2001
        If stretch_img = "NO" Then img_ctrl = "NO"    'march 31 2001
        sscreen_saver = "Y"
    End If
'    If UCase(Left(ttt, 3)) = "VVV" Then
'        ppaste = Right(ttt, Len(ttt) - 3)
'        GoTo What_50
 '   End If
    If UCase(Left(ttt, 2)) = "VV" Then
        Clipboard.SetText Right(ttt, Len(ttt) - 2)
        'see programmers guide page 170 re mdi forms and clipboard.settext
        GoTo What_50
    End If

    crlf = ""       'December 2 2000
    If UCase(ttt) = "QQQ" Then
        ttt = "RRR"
        crlf = "NO"
    End If
'add the option to do a search and replace
    If UCase(ttt) = "RRR" Then
        encript = "RRR"
'            xtemp = InputBox("27 October testing 2" + ttt, , , 4400, 4500) 'TESTING ONLY
        GoSub replace_29000
        GoTo File_40
    End If
'do the encription here november 20 2000   16Apr2014 need something like this max of 25000 lines for rewind bb option
line_70:
    save_line = "70"
    If UCase(ttt) <> "REVERSE" Then GoTo line_75
    ttt = ""
    FileExt = Cmd(19)  '16Apr2014a replace.txt
    Close OutFile       '16apr2014a
    DoEvents            '16apr2014a
    OutFile = FreeFile      'november 14 2000
    '        xtemp = InputBox("16Apr2014a reverse files (in file)= " + TheFile + " (out file) = " + FileExt, , , 4400, 4500) 'TESTING ONLY"
    '        If xtemp = "x" Then GoTo End_32000 '16apr2014 testing
    Open TheFile For Input As #OutFile  'the main file open for input here vip
    encript = "RRR"
    criptcnt = 0
line_73:
    Line Input #OutFile, aaa
    criptcnt = criptcnt + 1
    cript1(criptcnt) = aaa
    GoTo line_73
line_74:
    ExtFile = FreeFile      'november 14 2000
'    FileExt = "replace.txt" '16Apr2014b
    Open FileExt For Output As #ExtFile  '
    '        xtemp = InputBox("16Apr2014 testing total reversal array count " + CStr(criptcnt), , , 4400, 4500) 'TESTING ONLY
    '        If xtemp = "x" Then GoTo End_32000 '16apr2014 testing
    For i = criptcnt To 1 Step -1
        Print #ExtFile, cript1(i)
    Next i
    '        xtemp = InputBox("16Apr2014 testingb output done" + CStr(criptcnt), , , 4400, 4500) 'TESTING ONLY
    '        xtemp = InputBox("16Apr2014 OK reversal ranaming file " + FileExt + " as " + TheFile, , , 4400, 4500) 'TESTING ONLY
    '        Close ExtFile '16Apr2014b testing only
    '       If xtemp = "x" Then GoTo End_32000 '16apr2014 testing
    tt1 = UCase(InputBox("total reversals=" + CStr(criptcnt) + " rename files Y/N <Y>", "Rename Prompt", , xx1 - offset1, yy1 - offset2))
    If tt1 <> "N" Then
        tt1 = "Y"
    End If
    If UCase(tt1) = "N" Then
        Close ExtFile, OutFile
'        Close FileFile, OutFile
        GoTo line_74a
    End If
    '       xtemp = InputBox("16Apr2014 stepa " + FileExt + " as " + TheFile, , , 4400, 4500) 'TESTING ONLY
    Close ExtFile, OutFile
'    Close FileFile, OutFile
    '        xtemp = InputBox("16Apr2014 stepb " + FileExt + " as " + TheFile, , , 4400, 4500) 'TESTING ONLY
    DoEvents
'    Name TheFile As "c:\oldfile.txt"
'    Kill "oldfile.txt"  'december 3 2000
    DoEvents
    '        xtemp = InputBox("16Apr2014 stepc " + FileExt + " as " + TheFile, , , 4400, 4500) 'TESTING ONLY
'    Name TheFile As "oldfile.txt"
    DoEvents
    '        xtemp = InputBox("16Apr2014 stepd " + FileExt + " as " + TheFile, , , 4400, 4500) 'TESTING ONLY
    Kill TheFile
    DoEvents
    Name FileExt As TheFile
    'rename files here if Y entered
line_74a:
    save_line = "line_74a"
    GoTo End_32000
line_75:
    save_line = "75"
    
    If UCase(ttt) = "CRIPT" Then
        encript = "CRIPT"
        xtemp = Cmd(32)
        GoSub line_16000
        GoSub line_29500        'encription array read
        case_yes = "Y"
        GoSub replace_29000
        GoTo File_40
    End If
'do the de-encription here november 20 2000
    If UCase(ttt) = "DECRIPT" Then
        encript = "DECRIPT"
        xtemp = Cmd(32)
        GoSub line_16000
        GoSub line_29500        'encription array read
        case_yes = "Y"
        GoSub replace_29000
        GoTo File_40
    End If
'set up for the de-encription here november 20 2000
    If UCase(ttt) = "MYSTUF" Then
        encript = "MYSTUF"
        xtemp = Cmd(32)
        GoSub line_16000
        GoSub line_29500        'encription array read
        case_yes = "Y"
        GoTo What_50
    End If
    
'10 January 2005    If text_pause Then
    If text_pause And p2p2 <> "F" Then
        MAX_CNT = MAX_CNT + 1       '06 January 2005 no need to display the end of screen stuff when pausing
    End If              '06 January 2005   add 1 to the max number of lines to display here
' allow for the number of characters in string to do select
    If Len(ttt) = 0 Then GoTo line_80
    If Len(ttt) = 1 Then GoTo line_80
    If Len(ttt) = 2 And ttt = "P1" Then GoTo line_80
    If Len(ttt) = 2 And ttt = "CH" Then GoTo line_80
    If Len(ttt) = 2 Then
        ttt = "P"
        GoTo line_80
    End If
    If Len(ttt) = 3 Then
        ttt = "P1"
        GoTo line_80
    End If
    If Len(ttt) = 4 Then
        ttt = "CH"
        GoTo line_80
    End If
    If Len(ttt) = 5 Then
        ttt = "Z"
        GoTo line_80
    End If
    If Len(ttt) = 6 Then
        ttt = "E"
        GoTo line_80
    End If
    If Len(ttt) = 7 Then
        ttt = "F"
        GoTo line_80
    End If
    If Len(ttt) = 8 Then
        ttt = "S"
        GoTo line_80
    End If
    If Len(ttt) = 9 Then
        ttt = "X"
        GoTo line_80
    End If

line_80:
    Test1_str = "P"
 '       Print "Cmd(20)="; Cmd(20); "="
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 3", , , 4400, 4500)  'TESTING ONLY
    End If
    If ttt = "P1" Then
        ttt = "P"
        Test1_str = "P1"
    End If
    If ttt = "P2" Then
        ttt = "P"
        Test1_str = "P2"
    End If
    If ttt = "P3" Then
        ttt = "P"
        Test1_str = "P3"
    End If
    
    SAVE_ttt = ttt
'    frmproj2.Caption = program_info + " (cls #2)" '18Dec2013
    Cls     'clear screen each time
    'enter data to end of file
    Picture_Search = "NO"
    If ttt = "P" Then
        Picture_Search = "YES"
        ttt = "C"
        GoTo Do_Search_110
        SAVE_ttt = ttt
    End If                  'march 31/00
        
    If ttt = "Z" Then
        GoSub Do_Append_19000
        GoTo What_50
    End If                  'march 28/00
    If ttt = "CH" Then
        GoSub Do_Change_18000
        DoEvents
        GoTo What_50
    End If                  'march 29/00
    If ttt = "E" Then
        GoTo Do_Enter_20000
    End If
    'search and search with start
    If ttt = "S" Or ttt = "SS" Then
        GoTo Do_Search_110
    End If
    'flash search
    If ttt = "F" Then
'        tt1 = InputBox("testing flash prompt", , , 4400, 4500)  '04 September 2004
        auto_redraw = "NO"
        frmproj2.AutoRedraw = False     'november 11 2001
        'turn off redraw on flash display
        GoTo Do_Search_110
    End If
    
    'context search     aug 08/99
    If ttt = "C" Then
        GoTo Do_Search_110  'aug 08/99
    End If
    'quick search october 21 2000
    If ttt = "Q" Then
        GoTo Do_Search_110
    End If
    GoTo File_40
'    GoTo End_32000
    
Do_Search_110:
    save_line = "110"    'for error handling
    If img_ctrl = "YES" Then
        Set Image1.Picture = LoadPicture        'february 21 2001
    End If
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 3.1", , , 4400, 4500)  'TESTING ONLY
    End If
    
    

    Close OutFile
    OutFile = FreeFile      'november 14 2000
    If mbxyes = "Y" Then
        Open TheFile For Binary As #OutFile  'the main file open for input here vip
    Else
        Open TheFile For Input As #OutFile  'the main file open for input here vip
    End If
    SSS1 = ""
    SSS2 = ""
    SSS3 = ""
    SSS4 = ""       '09 JUNE 2002
    SSS5 = ""
    SSS6 = ""
    
'may 12 2001    dsp_cnt = 0
    zzz_cnt = 0
    zzz_chrs = 0
    tot_disp = 0        ' zero count of lines displayed
    end_cnt = 0         'allow for more than 1 eof in file

 '   ScrollBars = True
 '   Picture1.Picture = LoadPicture("c:\temp.bmp")
 '   Picture2.Picture = Picture1.Picture
 '   Picture2.Top = 0
 '   Picture2.Left = 0
'    Picture2.Width = Picture1.Width
 '   Picture2.Height = Picture1.Height
    
'        II = DoEvents()
    
 '   PaintPicture Picture1.Picture, -1, -3300, 9400, 8000
'      MyAppID = Shell("C:\WINDOWS\KODAKIMG.EXE", 1)

 
'    MyAppID = Shell("C:\WINDOWS\SYSTEM\VIEWERS\QUIKVIEW.EXE", 1)

' MyAppID = Shell("C:\PROGRAM FILES\ACCESSORIES\MSPAINT.EXE", 1)
'    AppActivate MyAppID
 '   SendKeys "^o", True
  '  SendKeys "c:\temp.tif", True
  '  SendKeys "{c:\temp.bmp}", True
    
'    SendKeys "%F1", 1
    
   ' SendKeys "%F0C:\TEMP.BMP", 1
  '  SendKeys " C:\TEMP.BMP"
    
  ' SendKeys Ctrl("o"), 1
  '  SendKeys "^co"
    
    
 '   MyAppID = Shell("C:\PROGRAM FILES\colordesk utilities\photo\cdphoto.EXE", 1)
    
  '  MyAppID = Shell("C:\fbscanner\imagein3\imagein.EXE", 1)
 '   MyAppID = Shell("C:\PRORAM FILES\ACCESSORIES\IMAGEIN.EXE", 1)
    
'*******************************************************
 '       * * *   S E A R C H   S T R I N G   * * * entry option prompt
'*******************************************************
'august 27/00
'    ttt = InputBox("Search String 'A' for ALL", , , 4400, 4500)
 '   ttt = UCase(ttt)
        TheSearch = "."
'28 april 2002    If sscreen_saver = "Y" Then
'         frmproj2.Caption = " do_search_110 dougdoug " + SAVE_SSS + "*" + interrupt_prompt2 + "*" + sscreen_saver + "*" + sscreen_saver_ww + "*" + SSS1 + "*" + SSS2 + "*" + ss_search '20 November 2004 test
    If sscreen_saver = "Y" And sscreen_saver_ww = "YES" And interrupt_prompt2 = "WW" Then   '21 November 2004
        ttt = SAVE_SSS      '21 November 2004  ie like "OLDIE/IRENE" IN SAVE_SSS
        delay_sec = new_delay_sec       '23 November 2004 set it back
        interrupt_prompt2 = ""          '23 November 2004 was not allowing a WW entry at prompt two
        GoTo line_600
    End If                          '21 November 2004
    If sscreen_saver = "Y" And sscreen_saver_ww <> "YES" Then
        ttt = ss_search         'screen saver logic
        GoTo line_600
    End If
line_500:               'december 3 2000
        If text_pause And inin <> "" Then
'    tt1 = InputBox("testing pause " + p2p2 + "*" + inin + "*", , , 4400, 4500)
            GoTo line_510    '05 october 2002
        End If
        If iimport = "Y" Then
            ttt = "kskdlskdj"   'skip search entry prompt
            GoTo line_510     'december 24 2000
        End If
    
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 3.2", , , 4400, 4500)  'TESTING ONLY
    End If
        
        GoSub Search_26000      'get the string to find search prompt here *******

line_510:               'december 24 2000
'    Print ttt   'testing
'    tt1 = InputBox("testing search_26000", , , 4400, 4500)

        
        If UCase(ttt) = "E" Or UCase(ttt) = "X" Then GoTo What_50
'        If Len(ttt) = 1 Then GoTo What_50   'january 05 2001
        If (prompt2 = "P1" Or prompt2 = "P") And _
            search_prompt = "in" And ttt = "" Then
            ttt = "PHOTO"
            GoTo line_600
        End If
        If (prompt2 = "P1" Or prompt2 = "P") And _
            UCase(ttt) = "A" Then
            ttt = "PHOTO"
            GoTo line_600
        End If
        If search_prompt = "in" And ttt = "" Then
            ttt = Cmd(33)       'december 7 2000
'            ttt = "D"
            GoTo line_600
        End If                  'april 24/00
                            'default to day search on start
        
    
    If UCase(ttt) = "A" Then GoTo line_600
    If UCase(ttt) = "D" Then GoTo line_600
    If UCase(ttt) = "M" Then GoTo line_600
'    If ttt = "MM" Then GoTo line_600
'    If ttt = "DD" Then GoTo line_600
'december 28 2000    If Len(ttt) = 1 Then ttt = "."   'force display old searches

'december 28 2000 comment out the logic below
'any 2 matching characters will do a paste function Ctrl/V
'    If Len(ttt) = 2 And Left(ttt, 1) = Right(ttt, 1) Then
'        ttt = Clipboard.GetText(vbCFText)
  '      Print ttt 'testing only
'    End If
line_600:
'    search_prompt = "D"
    search_prompt = Cmd(33) 'december 7 2000
     
    
    printed_cnt = 1         'may 10/00
    printed = "NO"
    
    TheSearch = ttt
'august 27/00
'    If TheSearch <> "" Then
'        GoSub Search_26000
 '   End If
    

    If ttt = "-" Then
        ttt = SAVE_SRCH + ""
    End If          'use last search if - entered 15mar00
    If ttt <> "" Then
        If UCase(ttt) <> "A" Then
        SAVE_SRCH = ttt + ""    'save search for reuse 15mar00
        End If
        'november 6 2000
        If prompt2 = "Q" And SSS <> "D" Then
            SAVE_SRCH = qqq + ""
        End If      'october 23 2000
    End If
    
    SSS = UCase(ttt)    'Ensure all is upper case for search
'november 6 2000
'december 8 2000    If prompt2 = "Q" And SSS <> "D" Then
    If prompt2 = "Q" And SSS <> "D" And uppercase = "N" Then
        SSS = qqq       'no upper case in quickie search october 23 2000
    End If
    SAVE_SSS = SSS + ""
    'check for tabs here later and make them 4 or 5 spaces
'20 August 2003    Cls     'june 27/99
    cnt = 0
    dblStart = Timer    'get the start time
    hi_lites = "NO"
    If SSS = "" Then
        Close #OutFile          'june 27/99
        II = DoEvents           'june 27/99
        GoTo What_50
    End If
'if D entered switch with current date ie 6/14/99 as the
' search string in the format 6/6/14/99
'06 December 2004 testing previous day.
'--------------------------------------------------
'date is as 12/6/2004 for December 06 2004 returned with time also
'this will not do a previous month ie if sitting on 01 good enought for now.
    If SSS = "DX" Then
        SSS = Format(Now, "ddddd ttttt")    'display today
        temp1 = InStr(2, SSS, " ")
        temptemp = Trim(Left(SSS, temp1))
        II = InStr(temptemp, "/")
        III = InStr(II + 1, temptemp, "/")
        temptemp = Mid(temptemp, II + 1, III - II + 1)
        temp1 = Val(temptemp) - 1       'the day is reduced by 1 here
        If temp1 = 0 Then temp1 = 1     'do not worry about last month for now
        temptemp = CStr(temp1)
        temptemp = Left(SSS, II) + CStr(Val(temptemp)) + Right(SSS, Len(SSS) - III + 1)
        SSS = temptemp
        temp1 = InStr(2, SSS, " ")  'case it changes from 10 to 9 etc
'            frmproj2.Caption = " 06 Dec 2004 testing=" + temptemp + "=" '06 December 2004
'            new_delay_sec = 2           '06 December 2004 testing
'            GoSub line_30300            '06 December 2004 testing
'        SSS = " " + sep + " " + sep + Left(SSS, temp1 - 1)
'09 june 2002        SSS = "M" + sep + ":" + sep + Left(SSS, temp1 - 1) 'january 05 2001
        SSS = "M" + sep + "M" + sep + "M" + sep + "M" + sep + ":" + sep + Left(SSS, temp1 - 1) 'january 05 2001
     End If

'--------------------------------------------------
    If SSS = "D" Then
        SSS = Format(Now, "ddddd ttttt")    'display today
'            frmproj2.Caption = " 06 Dec 2004 testing=" + SSS + "=" '06 December 2004
'            new_delay_sec = 2           '06 December 2004 testing
'            GoSub line_30300                        '06 December 2004 testing
        temp1 = InStr(2, SSS, " ")
'        SSS = " " + sep + " " + sep + Left(SSS, temp1 - 1)
'09 june 2002        SSS = "M" + sep + ":" + sep + Left(SSS, temp1 - 1) 'january 05 2001
        SSS = "M" + sep + "M" + sep + "M" + sep + "M" + sep + ":" + sep + Left(SSS, temp1 - 1) 'january 05 2001
     End If
        
    If SSS = "DD" Then
        SSS = " " + sep + " " + sep + date_displayed   'day of last find
     End If
     
    If SSS = "M" Then
        SSS = Format(Now, "ddddd ttttt")    'display this month
        temp1 = InStr(2, SSS, "/")  'end of month
        temp2 = InStr(2, SSS, " ")   'right after year
'        temp3 = InStr(temp1 + 1, SSS, "/") 'start of year
 '       SSS = " " + sep + Mid(SSS, temp3 + 1, temp2 - temp3) + sep + " " + Left(SSS, temp1)
            ' sep99 sep 7/ for july
        SSS1 = Left(SSS, temp1)
        SSS2 = Mid(SSS, temp2 - 3, 4)
        SSS3 = ""
        SSS4 = ""       '09 JUNE 2002
        SSS5 = ""
        SSS6 = ""
        
        SSS = ""
        KEEPS1 = SSS1 + ""
        KEEPS2 = SSS2 + ""
        KEEPS3 = SSS3 + ""
        KEEPS4 = SSS4 + ""      '09 june 2002
        KEEPS5 = SSS5 + ""
        KEEPS6 = SSS6 + ""
        
        SAVE_KEEPS1 = KEEPS1 + ""
        SAVE_KEEPS2 = KEEPS2 + ""
        SAVE_KEEPS3 = KEEPS3 + ""
        SAVE_KEEPS4 = KEEPS4 + ""       '09 june 2002
        SAVE_KEEPS5 = KEEPS5 + ""
        SAVE_KEEPS6 = KEEPS6 + ""
        inin = "M"
'maybe put the above changes in the MM below june 15/00
'        Context = "yes"
'        Print "testing="; SSS1; "*"; SSS2; "*"; SSS3; "*"
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 4", , , 4400, 4500)  'TESTING ONLY
    End If
        GoTo input_990
    End If
    
    If SSS = "MM" Then                  'month of last find
        SSS = date_displayed
        temp1 = InStr(2, SSS, "/")  'end of month
        temp2 = InStr(2, SSS, " ")   'right after year
        temp3 = InStr(temp1 + 1, SSS, "/") 'start of year
 '       SSS = " " + sep + Mid(SSS, temp3 + 1, temp2 - temp3) + sep + " " + Left(SSS, temp1)
            ' sep99 sep 7/ for july
        SSS1 = Left(SSS, temp1)
        SSS2 = Mid(SSS, temp2 - 3, 4)
        SSS3 = ""
        SSS4 = ""       '09 JUNE 2002
        SSS5 = ""
        SSS6 = ""
        SSS = ""
        KEEPS1 = SSS1 + ""
        KEEPS2 = SSS2 + ""
        KEEPS3 = SSS3 + ""
        KEEPS4 = SSS4 + ""      '09 june 2002
        KEEPS5 = SSS5 + ""
        KEEPS6 = SSS6 + ""
        SAVE_KEEPS1 = KEEPS1 + ""
        SAVE_KEEPS2 = KEEPS2 + ""
        SAVE_KEEPS3 = KEEPS3 + ""
        SAVE_KEEPS4 = KEEPS4 + ""   '09 june 2002
        SAVE_KEEPS5 = KEEPS5 + ""
        SAVE_KEEPS6 = KEEPS6 + ""
        inin = "MM"
        GoTo input_990
    End If
    

    Context = "no"      'june 26/99 what is with this line?
    'june 26/99
    If SSS = "-" Then
        Context = "yes"
        SSS = " " + sep + " " + sep + date_displayed
        SAVE_KEEPS1 = KEEPS1 + ""        'ALLOW FOR CONTEXT HILITE
        SAVE_KEEPS2 = KEEPS2 + ""          'june 26/99
        SAVE_KEEPS3 = KEEPS3 + ""
        SAVE_KEEPS4 = KEEPS4 + ""       '09 june 2002
        SAVE_KEEPS5 = KEEPS5 + ""
        SAVE_KEEPS6 = KEEPS6 + ""
    End If      'use date found from previous search to get
                'all data entered that day
    
    If SSS = "=" Then
        Context = "yes"
        SSS = date_displayed
        temp1 = InStr(SSS, "/")   'end of month
        temp2 = InStr(SSS, " ")    'right after year
        temp3 = InStr(temp1 + 1, SSS, "/") 'start of year
        SSS = " " + sep + Mid(SSS, temp3 + 1, temp2 - temp3) + sep + " " + Left(SSS, temp1)
            ' sep99 sep 7/ for july
        SAVE_KEEPS1 = KEEPS1 + ""        'ALLOW FOR CONTEXT HILITE
        SAVE_KEEPS2 = KEEPS2 + ""          'june 26/99
        SAVE_KEEPS3 = KEEPS3 + ""
        SAVE_KEEPS4 = KEEPS4 + ""
        SAVE_KEEPS5 = KEEPS5 + ""
        SAVE_KEEPS6 = KEEPS6 + ""
    End If
    
    'PARSE THE SSS STRING INTO PARTS SSS1 SSS2 SSS3 SSS4 SSS5 SSS6
line_700:
    i = InStr(SSS, "  ")
    If i <> 0 Then
        SSS = Left(SSS, i) + Mid(SSS, i + 2)
        GoTo line_700
    End If
    i = InStr(SSS, sep)
    SSS1 = SSS + ""
    s1len = Len(SSS1)
    inin = SSS + ""         'june 13/99

    If i = 0 Then
        GoTo input_990
    End If
    SSS1 = Left(SSS, i - 1)
    s1len = Len(SSS1)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

    i = InStr(SSS, sep)
    SSS2 = SSS
    s2len = Len(SSS2)
'        Print "testing="; SSS1; "*"; SSS2; "*"; SSS3; "*"
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 5", , , 4400, 4500)  'TESTING ONLY
    End If
'09 june 2002 add 3 more elements here
    If i = 0 Then
        GoTo input_990
    End If
    SSS2 = Left(SSS, i - 1)
    s2len = Len(SSS2)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

    i = InStr(SSS, sep)
    SSS3 = SSS
    s3len = Len(SSS3)
'---------------------------------------------
    If i = 0 Then
        GoTo input_990
    End If
    SSS3 = Left(SSS, i - 1)
    s3len = Len(SSS3)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

    i = InStr(SSS, sep)
    SSS4 = SSS
    s4len = Len(SSS4)
'========================================================
    If i = 0 Then
        GoTo input_990
    End If
    SSS4 = Left(SSS, i - 1)
    s4len = Len(SSS4)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

    i = InStr(SSS, sep)
    SSS5 = SSS
    s5len = Len(SSS5)

'----------------------------------------------
    If i = 0 Then
        GoTo input_990
    End If
    SSS5 = Left(SSS, i - 1)
    s5len = Len(SSS5)
    j = Len(SSS)
    SSS6 = Right(SSS, j - i)
    s6len = Len(SSS6)

input_990:

    If ss_only = "YES" And p2p2 = "SS" Then
        ttt = "SS" + " " + SSS1 + " " + SSS2 + " " + SSS3 + " " + SSS4 + " " + SSS5 + " " + SSS6
        ss_only = "NO"
        GoTo auto_p2
    End If                  '07 december 2002
    tot_cnt = 0
    GoSub line_14500        'check for imbedded spaces in search strings
                            'january 19 2001
        '***********************************************
        'Major input line for the sequential file read
        '***********************************************


input_1000:                                     'INPUT HERE
    save_line = "1000"    'for error handling
    slomo = False       '14 January 2004
    slomo = True        '31Dec2011      need this for the ability to interrupt ***vip*** maybe???
'14May2012
    cont_str = "S"      '14May2012 make it the default for each new set
    motion_yn = "NO"    '03 September 2004
    
    If SAVE_ttt = "C" Then
        Context_cnt = Context_cnt + 1   'aug 08/99
        If Context_cnt > MAX_CNT Then
            Context_cnt = 1
        End If
        Context_text(Context_cnt) = ccc + ""
    End If          'aug 08/99 save last 10 lines at least
    
    'the following code along with the form keypreview set to true
    'should enable escape similar to ctrl/c on the vax
    'endscript and keydown subroutines at the end also Dec 02/99
    If zzz_cnt Mod 1000 = 0 Then
        DoEvents
    End If
    
    If Escape Then
        Escape = False
        GoTo End_32000
    End If
    If Picture_Search = "YES" Then
        Previous_line = ooo     'keep this to search for XXX.
    End If                  'march 31/00
    'main input file read here
    
       '*************************************************
       'main input file data read             ' * * * I N P U T * * * INPUT INPUT INPUT
       '*************************************************
input_1000a:
    
    Line Input #OutFile, aaa
            
'29Oct2012            If mixx = True Then
            If mixx = True And rand <> True Then
                If Left(UCase(aaa), 5) <> "PHOTO" Then
                    ss_search = "A"
                    mpg_file = ""  'change from no to blank
                    Picture_Search = "NO"
                    sscreen_saver_ww = "NO"
                    xxx_found = ""
                    interrupt_prompt2 = ""
                    inin = "A"  'tested and needed
                    p2p2 = "C"
                    ttt = "C"
                    text_pause = True '29Jul2012
'testing now                    Test2_str = "A"
                    sscreen_saver = "N"  'text will not show without this?
                Else
                    Test1_str = "P1" '13Jul2012  this seems very important very important
                    p2p2 = "WW"     '13Jul2012
                    Line_Search = "" '13Jul2012
                    sscreen_saver = "Y" '13Jul2012
                    ss_search = "photo"
'11Jul2012 test comment this out                    mpg_file = "YES"
                    Picture_Search = "YES"
                    sscreen_saver_ww = "YES"
                    xxx_found = ""
                    interrupt_prompt2 = ""
                    inin = "PHOTO"      '13Jul2012
                    ttt = "SS"
'                    Test2_str = ""
                    sscreen_saver = "Y"
'    frmproj2.Caption = program_info + " (cls #3)" '18Dec2013
                                        Cls     '12Nov2012
                End If
            End If              '30Jun2012b
'    If UCase(Left(aaa, 5)) = "PHOTO" Then
'            frmproj2.Caption = App.EXEName + " 30Jun2012a -- aaa =" + aaa + " * " + ss_search + " * " + mpg_file + " * " + Picture_Search + " * " + sscreen_saver_ww + " * " + xxx_found
'            ggg = InputBox(" 30Jun2012a (ss_search) (mpg_file) (picture_search) (sscreen_saver_ww) (xxx_found)", , , 4400, 4500)  'TESTING ONLY
'            frmproj2.Caption = App.EXEName + " 30Jun2012aa -- " + p2p2 + " * " + Line_Search + " * " + sscreen_saver + " * " + inin + " * " + Test1_str
'            ggg = InputBox(" 30Jun2012aa (p2p2) (line_search) (sscreen_saver) (inin) (Test1_str) ", , , 4400, 4500)  'TESTING ONLY
'            If ggg = "x" Then GoTo End_32000  '30Jun2012
'    End If           '30Jun2012
'22Jun2012
'If Left(UCase(App.EXEName), 7) = "BIGTEXT" Then
'    ooo = InputBox("", "data to display", , xx1 - ppoffset1, yy1 - ppoffset2) '
'    If ooo = "" Then GoTo End_32000
'        'ooo = "test this Print and whatever else we need to do right here"
'        aaa = ooo
'        vvv = ooo
'End If    '22Jun2012
'            frmproj2.Caption = " testing=" + Left(aaa, 40) + "=" '26 November 2004
'20 November 2004 maybe the line below needs setting as I return here a lot... ***vip*** todo check out sometime
'    save_line = "1000"    'for error handling 20 November 2004
'    line_pos = 0        'november 22 2000 january 05 2001
    zzz_cnt = zzz_cnt + 1
'29Oct2012
        If rand_prog And mixx And rand <> True And Left(UCase(Trim(aaa)), 5) = "BREAK" Then rand = True  '29Oct2012
'        If rand Then frmproj2.Caption = "Watching randomizer rand_str=" + rand_str + " rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num)
'       xtemp = InputBox("12Nov2012b (need a random number here) break_num zzz_cnt " + CStr(break_num) + " " + CStr(zzz_cnt) + " " + Left(aaa, 10), , , 4400, 4500) 'TESTING ONLY
    If rand Then
        '27Aug2010 maybe save the last break line here? below is the first "BREAK" line before the match line
        If Left(Cmd(81), 10) = "RAND_GROUP" And Left(UCase(Trim(aaa)), 5) = "BREAK" Then break_num = zzz_cnt '27Aug2010
'12May2012 test rand below testing randomizer testing randomizer
'        frmproj2.Caption = "11Nov2012a testing randomizer rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num) + " aaa=" + Left(aaa, 20)
'        xtemp = InputBox(" 11Nov2012a testing doug randomizer rnd  " + CStr(rand_no) + " " + CStr(zzz_cnt), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
       If zzz_cnt < rand_no Then GoTo input_1000a '18 august 2002
'        frmproj2.Caption = "11Nov2012ab testing randomizer rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num) + " aaa=" + Left(aaa, 20)
        '03 August 2003 below is where to change the 2 to what ever the range is for random by group display
        'ie if there are 6 videos in the group the number should be 12 or 13 below???  ***vip*** todo
    '06 September 2004 if random and a text search just do random on first get then sequential
'12Nov2012        If p2p2 = "C" Or p2p2 = "S" Or p2p2 = "F" Then
        If mixx <> True And (p2p2 = "C" Or p2p2 = "S" Or p2p2 = "F") Then
            rand = False
            GoTo input_1000a
        End If              '06 September 2004
'        frmproj2.Caption = "11Nov2012ac testing randomizer rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num) + " aaa=" + Left(aaa, 20)
'27Aug2010 skip out on the next break line after a match is found. so we need one at the end of the file Maybe?
      '03Nov2012 does not seem to be generating another number here
'12Nov2012 change this logic        If (Left(Cmd(81), 10) = "RAND_GROUP" And Left(UCase(Trim(aaa)), 5) = "BREAK") Or Left(Cmd(81), 10) <> "RAND_GROUP" Then  '27Aug2010
        If group_match = True And Left(UCase(Trim(aaa)), 5) <> "BREAK" And zzz_cnt < rand_no + 3 Then GoTo input_1000az '12Nov2012
        frmproj2.Caption = "showing randomizer rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num) + " aaa=" + Left(aaa, 20)
'        If (Left(Cmd(81), 10) = "RAND_GROUP" And Left(UCase(Trim(aaa)), 5) <> "BREAK") And break_num <> 0 And group_match = True Then GoTo input_1000az    '12Nov2012
'25Mar2014 maybe here
'        If zzz_cnt > rand_no + 2 Then    'this may display 2 in a row but that is ok  27Aug2010 random group needs to look for next BREAK line
        If zzz_cnt > rand_no + 2 And pauseyn <> "Y" Then    '25Mar2014
            'refresh, #outfile need to be able to reset the file at beginning
            Close #OutFile
            II = DoEvents
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            II = DoEvents       'yield to operating system
'29Oct2012 watch randomizer here
            
            rand_no = Int(rand_cnt * Rnd + 1)   'get a new random number
            rand_str = "random#" + CStr(rand_no) '18Mar2012
        frmproj2.Caption = "showing randomizer a rand_str=" + rand_str + " rnd=" + CStr(rand_no) + " " + CStr(zzz_cnt) + " break_num=" + CStr(break_num)
'       xtemp = InputBox("12Nov2012b (need a random number here) break_num zzz_cnt " + CStr(break_num) + " " + CStr(zzz_cnt) + " " + Left(aaa, 10), , , 4400, 4500) 'TESTING ONLY
            zzz_cnt = 0
            break_num = 0           '27Aug2010   from 999999???
            group_match = 0     '29Oct2012
            GoTo input_1000a
        End If
input_1000az:       '12Nov2012
        If Left(UCase(Trim(aaa)), 5) = "BREAK" Then GoTo input_1000a '14Nov2012 all this line is for is a BREAK was showing up
'        frmproj2.Caption = "11Nov2012ae testing randomizer rnd=" + CStr(rand_no) + " zzz_cnt=" + CStr(zzz_cnt) + " break_num=" + CStr(break_num) + " aaa=" + Left(aaa, 20)
'12Nov2012 change this logic        End If      '27Aug2010 should keep going till "BREAK" found if "RAND_GROUP"
    End If
'12May2012 test rand below
'        frmproj2.Caption = "testing doug randomizer rnd #3 " + aaa + SSS1 + " " + CStr(rand_no) + " " + CStr(zzz_cnt)
'       xtemp = InputBox("looking for match ( in 2 lines) break_num zzz_cnt " + CStr(break_num) + " " + CStr(zzz_cnt), , , 4400, 4500)  'TESTING ONLY
'    If hi_lites <> "YES" Then line_match = ""     'november 21 2000
    If encript = "MYSTUF" Then
        GoSub line_29400
'        If zzz_cnt Mod 10000 = 0 Then
'            DoEvents
'            Print "searching "; zzz_cnt
'        End If
    End If
    If skip_info <> "" Then
        If InStr(aaa, skip_info) <> 0 Then
            GoTo input_1000a
        End If
    End If          'november 14 2000  this check takes 1/2 second on 190,000 records
                    'ie 9.67 to 10.20 in the quickie mode for the if skip_info <> "" 190,000 times
'12May2012 test rand below
 '        frmproj2.Caption = "testing doug randomizer rnd #4 " + aaa + " " + SSS1 + " " + CStr(rand_no) + " " + CStr(zzz_cnt)
'       xtemp = InputBox(" 11Nov2012(match found *****) break_num zzz_cnt prompt2=" + prompt2 + CStr(break_num) + " " + CStr(zzz_cnt), , , 4400, 4500) 'TESTING ONLY
   If prompt2 <> "Q" Then
        GoTo input_1000b    'Quickie search fix
    End If
    

'    aaa = UCase(aaa)    'december 8 testing ucased timing
'info in testing the above line it took over 2 times as much
'time to do the search with just including the above statement'
'my outmail.txt search went from 14.45 seconds to 34.45 secs
    If uppercase = "Y" Then
         ooo = aaa + ""
         aaa = UCase(aaa) 'december 8 2000
    End If
    
    If InStr(aaa, SSS1) = 0 Then
        GoTo input_1000a
    End If
    If SSS2 = "" Then
        GoTo input_1000aa
    End If
    If InStr(aaa, SSS2) = 0 Then
        GoTo input_1000a
    End If
    If SSS3 = "" Then
        GoTo input_1000aa
    End If
    If InStr(aaa, SSS3) = 0 Then
        GoTo input_1000a
    End If

'add in sss4 thru sss6   09 june 2002
    If SSS4 = "" Then
        GoTo input_1000aa
    End If
    If InStr(aaa, SSS4) = 0 Then
        GoTo input_1000a
    End If
    
    If SSS5 = "" Then
        GoTo input_1000aa
    End If
    If InStr(aaa, SSS5) = 0 Then
        GoTo input_1000a
    End If

    If SSS6 = "" Then
        GoTo input_1000aa
    End If
    If InStr(aaa, SSS6) = 0 Then
        GoTo input_1000a
    End If

'    If InStr(aaa, qqq) = 0 Then
'        GoTo input_1000a
'    End If      'october 21 2000
input_1000aa:
    If uppercase = "Y" Then aaa = ooo 'december 8 2000
        previous_count = previous_count + 1   'october 22 2000
        If previous_count > 100 Then
            previous_count = 1
        End If
        previous_picture(previous_count) = zzz_cnt

input_1000b:
'12May2012 test rand below
'14May2012 aaa = UCase(aaa)    '12May2012
'        frmproj2.Caption = "testing 12May2012 randomizer rnd #5 " + aaa + " " + SSS1 + " " + CStr(rand_no) + " " + CStr(zzz_cnt)
'       xtemp = InputBox(" (match found *****) break_num zzz_cnt prompt2=" + prompt2 + CStr(break_num) + " " + CStr(zzz_cnt), , , 4400, 4500) 'TESTING ONLY
    'december 11 2000 the lines from here to input_1000bb
'december 15 2000 testing stuff below
'
'       xtemp = InputBox(" 29Oct2012a(match found *****) break_num zzz_cnt prompt2=" + prompt2 + CStr(break_num) + " " + CStr(zzz_cnt) + " aaa=" + Left(aaa, 20), , , 4400, 4500) 'TESTING ONLY
      '23Mar2014 might need a change here for fast forward where there is a pause in the search string
            If pauseyn = "Y" And rand And UCase(Left(Cmd(81), 10)) = "RAND_GROUP" Then group_match = True '23Mar2014
            If mixx And rand And UCase(Left(Cmd(81), 10)) = "RAND_GROUP" Then group_match = True     '29Oct2012
            If show_files_yn Then       '24 december 2002
'       xtemp = InputBox("DOUG " + show_files+" "+aaa, , , 4400, 4500) 'TESTING ONLY
        II = InStr(aaa, " append start") 'other appends have -append start
        If II <> 0 Then
            show_files = Left(aaa, II - 1)
        End If
    End If                      '24 december 2002
        
        xtemp = ""
'        xtemp = "test="
If xtemp <> "" Then
    III = Len(aaa)      'december 17 2000
    If III > 10 Then III = 10
        For II = 1 To III
        xtemp = xtemp + CStr(Asc(Mid(aaa, II, 1))) + " "
        Next II
        aaa = aaa + xtemp + CStr(Len(aaa))
End If      'end of test section
    
    
    If emailsea <> "Y" Then GoTo line_1001
    If Right(aaa, 3) = "=20" Then
        aaa = Left(aaa, Len(aaa) - 3)
    End If          'december 18 2000 whatever the deal is with the "-20 " get rid of it
    'december 17 2000 if more than 1 ascii value < 10 then mbx start of new mail message
    If mbxyes <> "Y" Then GoTo nombx1000    'december 17 2000
    If Len(aaa) = 0 And mbxi = 0 Then GoTo input_1000a   'december 18 2000
    If mbxi = 0 Then III = InStr(aaa, "From:") 'december 17 2000
    If mbxi = 0 And III <> 0 Then
        aaa = Right(aaa, Len(aaa) - III + 1)
        dateskip = ""           'december 19 20002
        mbxi = 1
    End If
    
    If mbxi = 0 Then GoTo input_1000a 'december 18 2000

nombx1000:
        

    'december 15 2000 added and in test mode
    If InStr(aaa, "Reply-To") <> 0 And mbxyes = "Y" Then
        dateskip = ""
        boundarystr = ""
        boundarycnt = 0
        aaa = "========================= email start ========================="
        GoTo input_1000bb   'need to print the seperator above
    End If
    If InStr(aaa, "Return-Path") <> 0 Then
            GoTo input_1000a
    End If                  'december 20 2000
    If InStr(aaa, "Return-Path") <> 0 Then
        dateskip = ""
        boundarystr = ""
        boundarycnt = 0
        aaa = "========================= email start ========================="
        GoTo input_1000bb   'need to print the seperator above
    End If
    If InStr(aaa, "From - ") <> 0 Then
        dateskip = "Y"
        boundarystr = ""
        boundarycnt = 0
        aaa = "========================= email start ========================="
        GoTo input_1000bb   'need to print the seperator above
    End If
    If InStr(aaa, "From ????") <> 0 Then
        dateskip = "Y"
        boundarystr = ""
        boundarycnt = 0
        aaa = "========================= email start ========================="
        GoTo input_1000bb   'need to print the seperator above
    End If
    If dateskip = "F" Then
        dateskip = "Y"
        boundarystr = ""
        boundarycnt = 0
        aaa = "========================= email start ========================="
        GoTo input_1000bb   'need to print the seperator above
    End If
    If InStr(aaa, "Date: ") <> 0 And dateskip = "Y" Then
        dateskip = ""
        GoTo input_1000bb
    End If
    If InStr(aaa, "From: ") <> 0 And dateskip = "Y" Then
        dateskip = ""
        GoTo input_1000bb
    End If
    If dateskip = "Y" Then GoTo input_1000a     'december 12 2000
    If InStr(aaa, "Reply-To") <> 0 Then GoTo input_1000a
    If InStr(aaa, "To: ") <> 0 Then GoTo input_1000bb
    If InStr(aaa, "Subject: ") <> 0 Then GoTo input_1000bb
'    If InStr(aaa, "boundary=") <> 0 Then february 28 2001
    If InStr(aaa, "boundary=""") <> 0 Then
        III = InStr(aaa, "boundary=")
        II = InStr(III + 10, aaa, """") 'quote mark search double quote
        boundarystr = Mid(aaa, III + 10, II - III - 10)
        boundarycnt = 0     'december 18 2000
 '       tt1 = InputBox(boundarystr + " " + CStr(boundarycnt), , , 4400, 4500) 'TESTING ONLY
        GoTo input_1000a
    End If
    If Len(boundarystr) > 2 And InStr(aaa, boundarystr) <> 0 Then
        boundarycnt = boundarycnt + 1
        GoTo input_1000a
    End If
    If boundarycnt = 0 And Len(boundarystr) > 2 Then GoTo input_1000a
    If boundarycnt > 1 Then GoTo input_1000a
'december 30 2000 skip a few of the odd characters that linger
If mbxyes = "Y" And Len(aaa) < 6 Then
    III = Len(aaa)      'december 17 2000
        For II = 1 To III
        If Asc(Mid(aaa, II, 1)) > 126 Then GoTo input_1000a
        If Asc(Mid(aaa, II, 1)) < 9 Then GoTo input_1000a
        Next II
End If      'december 30 2000 end of code
    
input_1000bb:
    If InStr(aaa, "= email start =") <> 0 Then tot_s1 = tot_s1 + 1    'january 28 2001
    If InStr(aaa, "Errors-to:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Mime-Version:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Reply-to:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "User-Agent:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-Accept-Language:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Importance:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-Mozilla") <> 0 Then GoTo input_1000a
    If InStr(aaa, "References:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-Mailer:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-MSMail-") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-MimeOLE") <> 0 Then GoTo input_1000a
    If InStr(aaa, "X-Priority:") <> 0 Then GoTo input_1000a
    If InStr(aaa, "MIME-V") <> 0 Then GoTo input_1000a
    If InStr(aaa, "MIME format.") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Message-ID") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Message-Id: ") <> 0 Then GoTo input_1000a
    If InStr(aaa, "Content-T") <> 0 Then GoTo input_1000a
    If InStr(aaa, "charset=") <> 0 Then GoTo input_1000a
    'the characters "photo" must exist for this to match up
line_1001:      'december 25/2000
    If Picture_Search = "YES" Then '
        If InStr(UCase(aaa), "PHOTO") = 0 Then
        'don't skip if search is for XXX.
        ooo = aaa + ""
        GoTo input_1000
         End If
    End If          'march 31/00
'25 July 2003    If temp_sec <> -1 And temp_sec <> delay_sec Then
    If temp_sec <> 0 And temp_sec <> -1 And temp_sec <> delay_sec Then
        delay_sec = temp_sec
'        If delay_sec < 0.3 Then
'            tt1 = InputBox("testing=" + Format(delay_sec, "###0.000") + "=", , , 4400, 4500) 'TESTING ONLY
'       End If          '25 July 2003
    End If          'march 14 2001
    If Picture_Search = "YES" Then 'start of picture if statement----------------------------------------------
        mpg_file = "NO"         '06Apr2012
        III = InStr(UCase(aaa), " WAIT=")
        line_delay_sec = 0              '19 July 2003 Ver=1.07T
    '22 september 2003    If III <> 0 Then
'        If III = 0 And thumb_nail <> "YES" Then 21Jan2012 deactivate this for now and in future only use it for my laptop maybe
        If III = 99 And thumb_nail <> "YES" Then
            aaa = aaa + " WAIT=6.5 " + CStr(video_length)     '19Jan2012  Maybe need to set this real low and when it starts move it to a much larger number
            III = InStr(UCase(aaa), " WAIT=")  'as long as the time is less than the video length it will start ok
        End If              '19Jan2012   force the wait it seems to work on my laptop with wait
        If III <> 0 And thumb_nail <> "YES" Then
            temp_sec = delay_sec
            II = InStr(III + 5, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 6, II - III - 6)
'            tt1 = InputBox("testing=" + xtemp + "=", , , 4400, 4500) 'TESTING ONLY 08Nov2011doug
            delay_sec = Val(xtemp)
            line_delay_sec = delay_sec                  '19 July 2003 Ver=1.07T
        End If
        keep_line_delay_sec = line_delay_sec        '08Nov2011
'16 November 2003

        III = InStr(UCase(aaa), " COLOR=")      '29Sep2016
        If III = 0 Then GoTo line_1001a         '29Sep2016
        If Mid(UCase(aaa), III + 7, 3) = "BLA" Then Def_Fore = 0    'BLACK
        If Mid(UCase(aaa), III + 7, 3) = "NAV" Then Def_Fore = 1    'NAVY BLUE
        If Mid(UCase(aaa), III + 7, 3) = "GRE" Then Def_Fore = 2    'GREEN
        If Mid(UCase(aaa), III + 7, 3) = "AQU" Then Def_Fore = 3    'AQUA
        If Mid(UCase(aaa), III + 7, 3) = "BRO" Then Def_Fore = 4    'BROWN
        If Mid(UCase(aaa), III + 7, 3) = "PUR" Then Def_Fore = 5    'PURPLE
        If Mid(UCase(aaa), III + 7, 3) = "LIM" Then Def_Fore = 6    'LIME GREEN
        If Mid(UCase(aaa), III + 7, 3) = "GRE" Then Def_Fore = 7    'GREY
        If Mid(UCase(aaa), III + 7, 3) = "DAR" Then Def_Fore = 8    'DARK GREY
        If Mid(UCase(aaa), III + 7, 3) = "BLU" Then Def_Fore = 9    'BLUE BRIGHT
        If Mid(UCase(aaa), III + 7, 3) = "LGB" Then Def_Fore = 10    'LIME GREEN BRIGHT
        If Mid(UCase(aaa), III + 7, 2) = "PB" Then Def_Fore = 11    'PALE BLUE
        If Mid(UCase(aaa), III + 7, 3) = "RED" Then Def_Fore = 12    'RED
        If Mid(UCase(aaa), III + 7, 3) = "PIN" Then Def_Fore = 13    'PINK
        If Mid(UCase(aaa), III + 7, 3) = "YEL" Then Def_Fore = 14    'YELLOW
        If Mid(UCase(aaa), III + 7, 3) = "WHI" Then Def_Fore = 15    'WHITE
        Hold_Fore = Def_Fore        '12Feb2017
        
line_1001a:     '29Sep2016

        III = InStr(UCase(aaa), " SPEED=")
        line_speed = 1000
        play_speed = line_speed     '13 May 2004 (see other setting too) 28Apr2010 reactivating this line fixed the slow motion on non applicable ones
        If III <> 0 Then
'            temp_sec = delay_sec
            II = InStr(III + 6, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 7, II - III - 7)
'            tt1 = InputBox("testing=" + xtemp + "=", , , 4400, 4500) 'TESTING ONLY
            line_speed = Val(xtemp)
            play_speed = line_speed
        End If      '16 November 2003
        keep_play_speed = play_speed            '08Nov2011
'        keep_slomo = False                      '14Nov2011
        If keep_play_speed <> 1000 Then keep_slomo = True   '08Nov2011??
        slomo = True            '31Dec2011 somehow speed makes a big difference in the again stuff???
        '25Dec2011 also see the replay_yn logic how does that work
'24 September 2003 add the line_freeze_sec stuff
        III = InStr(UCase(aaa), " FREEZE=")
        Line_freeze_sec = 0
        If III <> 0 Then
            temp_sec = delay_sec
            II = InStr(III + 7, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 8, II - III - 8)
'            tt1 = InputBox("testing=" + xtemp + "=", , , 4400, 4500) 'TESTING ONLY
            Line_freeze_sec = Val(xtemp)
        End If

'24 September 2003 end of line_freeze_sec stuff
        
'22 March 2004        line_start_point = 0                            '19 July 2003 Ver=1.07T
        If thumb_nail <> "YES" Then line_start_point = 0                            '22 March 2004
        If thumb_nail <> "YES" Then line_start_point = 10                            '11 March 2007
        III = InStr(UCase(aaa), " START==")             '19 July 2003 Ver=1.07T
        line_start_point = 10        '28Dec2011
 '       If keep_line_start_point <> 0 Then line_start_point = keep_line_start_point  '25Dec2011
        If III <> 0 Then
            II = InStr(III + 7, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 8, II - III - 8)
'02Nov2011            line_start_point = Val(xtemp)
            line_start_point = Val(xtemp) + (OffSet * 1000)      '02Nov2011
            If line_start_point < 0 Then line_start_point = 10   '02Nov2011
' testing = "mpg_file11.6=" + mpg_file + " " + CStr(line_start_point) + " " + Left(mssg, 10) + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
        End If
       keep_line_start_point = line_start_point            '28Dec2011 this is important placement???
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
'      Test2_str = InputBox("E or X to Stop --  Enter . to rewind video (a) for all " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , xx1, yy1) '24 February 2004
'      Test2_str = InputBox("testing keep_line_start_point only " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , xx1, yy1) '24 February 2004
'        begin = ""          '15Sep2011
        begin = "NO"          '25Dec2011
'        begin_point = 0       '08Nov2011
        begin_point = line_start_point          '16Dec2011 moved from below 28Dec2011
        III = InStr(UCase(aaa), " BEGIN")           '15Sep2011
        If III > 0 Then
                begin = "YES"                   '15Sep2011
                begin_point = 10                                 '20Sep2011 & 08Nov2011? change from 0 to 10
        End If
        III = InStr(UCase(aaa), " BEGIN==")             '20Sep2011
        If III <> 0 Then
            II = InStr(III + 7, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 8, II - III - 8)
'02Nov2011            begin_point = Val(xtemp)
            begin_point = Val(xtemp) + (OffSet * 1000)       '02Nov2011
        End If                                          '20Sep2011
        If begin_point < 0 Then
            begin_point = line_start_point + begin_point    '27Sep2011
            If begin_point < 0 Then begin_point = 10        '27Sep2011
        End If                                              '27Sep2011
    If altt Then                                   '12Aug2012
        III = InStr(UCase(aaa), " ALT==")             '12Aug2012
        If III <> 0 Then
            II = InStr(III + 5, aaa + " ", " ") 'make sure there is a trailing space here
            xtemp = Mid(aaa, III + 6, II - III - 6)
            alt_amt = Val(xtemp) + (OffSet * 1000)       '12Aug2012
            begin_point = begin_point + alt_amt     '12Aug2012
            keep_line_start_point = keep_line_start_point + alt_amt '12Aug2012
            line_start_point = line_start_point + alt_amt '12Aug2012
        End If                                          '12Aug2012
    End If                                      '12Aug2012
'30Jan2012
' testtest = "delay_sec=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo)  '17Jan2012 teststring
'       testtest = InputBox("check delay_sec info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'        If line_start_point = 10 And begin = "NO" And delay_sec = 0 Then
'19Jan2012 The logic below makes it run a 2nd time when there is no timing controls it is needed....
' for some reason my PC works fine now with this logic for the "G" stuff but not my laptop thus the changes for today???
' hopefully they will solve the problem on the laptop. If I hit enter at the right time the laptop version works????
' also if there is a "wait=" in the control file my laptop works so I will try forcing that if thumb_nail <> "YES" and mpg_file = "YES"
'    If thumb_nail <> "YES" And mpg_file = "YES" And begin = "NOoo" And line_start_point = 10 Then
        'III = InStr(UCase(aaa), " WAIT=")
'        III = 0        'force it for now
'        line_delay_sec = 0
'        Beep 'just to let me know
'        If III = 0 Then
'            temp_sec = delay_sec
'            delay_sec = 1111
'            line_delay_sec = delay_sec
'            keep_line_delay_sec = line_delay_sec
'         End If
'   End If              '19Jan2012   first stab at trying to get by without wait= for my laptop
'09Feb2012        If line_start_point = 10 And begin = "NO" And delay_sec = Val(Cmd(27)) Then
'needed to add the mpg_file stuff below because ww for pictures was getting the huge delay values of 11111
        If line_start_point = 10 And begin = "NO" And delay_sec = Val(Cmd(27)) And mpg_file = "YES" Then
                DoEvents
                DoEvents
                begin = "YES"
                begin_point = 10
        '        line_start_point = 10
        '        keep_line_start_point = 10
        '        play_speed = 777
        '        keep_play_speed = 777
                keep_begin = begin
                keep_begin_point = begin_point
        '16Jan2012 and check above for line_delay_sec        delay_sec = 10000
'17Jan2012
'        frmproj2.Caption = "(Finding Video length)"
 'i = mciSendString("status video1 length wait", mssg, 255, 0)
' DoEvents
 'temp3 = InStr(mssg, Chr$(0))
' video_length = Val(Left(mssg, temp3 - 1))
 '17Jan2012 maybe the file isn't opened yet
' testtest = "delay_sec=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo)  '17Jan2012 teststring
'       testtest = InputBox("check delay_sec info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
               
'                delay_sec = video_length / 1000                '16Jan2012 this might need to change...
'09Feb2012 the delay_sec below was causing a problem till i checked mpg_file above
                delay_sec = 11111    ' 17Jan2012 if no begin and no start play the full video on "G" entry
                line_delay_sec = delay_sec
                keep_line_delay_sec = delay_sec
                new_delay_sec = delay_sec
                slomo = True
                keep_slomo = True
        End If                                  '12Jan2012 when no start or begin set begin point
        keep_begin_point = begin_point          '08Nov2011
        keep_begin = begin                      '08Nov2011?
'        keep_line_start_point = begin_point     '08Nov2011
        resume_str = ""                 '25Sep2011
        III = InStr(UCase(aaa), " RESUME")             '25Sep2011
        If III > 0 Then resume_str = "YES"              '25Sep2011
        keep_resume_str = resume_str                '08Nov2011
        again_str = ""                  '08Nov2011
        III = InStr(UCase(aaa), " AGAIN")               '08Nov2011
        If III > 0 Then again_str = "YES"               '08Nov2011
        
        linebudgeyn = False                             '17Dec2017
        totbud = 0                                      '06Jan2018
        smlbud = Val(Cmd(85))                           '17Dec2017
        III = InStr(UCase(aaa), "SMLBUD")                '17DEC2017
'            frmproj2.Caption = LCase(temptemp) + " linebudge***** " + budge_str + " " + tempdata + " of " + CStr(video_length) '06Jan2018
'        temptemp = InputBox("11a January 2018 test ", aaa + " budge " + CStr(JJ) + " " + CStr(III) + " ", , xx1 - 5000, yy1 - 5000) '07Jan2018
        If III = 0 Then
            GoTo line_1001b                 '11Jan2018
        End If
        linebudgeyn = True               '17Dec2017
'            frmproj2.Caption = LCase(temptemp) + " linebudge********* " + budge_str + " " + tempdata + " of " + CStr(video_length) '06Jan2018
'        temptemp = InputBox("11b January 2018 test ", aaa + " budge " + CStr(JJ) + " " + CStr(III) + " ", , xx1 - 5000, yy1 - 5000) '07Jan2018
        JJ = InStr(III + 6, aaa, " ")
        If JJ = 0 Then GoTo line_1001b
        If JJ > III + 6 Then smlbud = Val(Mid(aaa, III + 6, JJ - III + 6))  '17Dec2017
        If smlbud = 0 And linebudgeyn = True Then smlbud = 2        'budge with no budge amount crazy
        If linebudgeyn = True Then budge_str = "YES"        '17Dec2017

line_1001b:
'2018
    save_line = "1001b"      'for error handling
        III = InStr(UCase(aaa), "BIGBUD")                '17DEC2017
        If III = 0 Then GoTo line_1001c
        linebudgeyn = True               '17Dec2017
        JJ = InStr(III + 6, aaa, " ")
        If JJ = III + 6 Then GoTo line_1001c
        If JJ > III + 6 Then bigbud = Val(Mid(aaa, III + 6, JJ - III + 6))  '17Dec2017
        If bigbud = 0 And linebudgeyn = True Then bigbud = 300        'budge with no budge amount crazy
        budge_str = ""          '17Dec2017
        If linebudgeyn = True Then budge_str = "YES"        '17Dec2017
         
line_1001c:
        continueyn = "N"                            '28Jan2012
        III = InStr(UCase(aaa), " CONTINUE")        '28Jan2012
        If III > 0 Then continueyn = "Y"            '28Jan2012
        pauseyn = "N"                            '28Jan2012
        III = InStr(UCase(aaa), " PAUSE")        '28Jan2012
        If III > 0 Then pauseyn = "Y"            '28Jan2012
        hold_pauseyn = pauseyn                  '25Mar2012
'   If again_str = "YES" Then slomo = False          '14Nov2011
'for the AGAIN stuff to work for fts objects the speed can only be less than 1000 ie 999 till this is fixed???
'it works for speed=200 or such but not 1000 need to eventually fix this but for now 999 is good enough
    End If          'march 14 2001 end of picture if statement---------------------------------------------
        
    'replace any tabs with 4 spaces right here
    save_line = "1002"      'for error handling
line_1002:
'07 november 2002 comment out the do_tab below for now
'    If do_tab Then      '05 october 2002 (no spaces in search string)
    tt = InStr(aaa, Chr(9)) 'check for tabs
    If tt = 0 Then
        GoTo line_1008
    End If
    'change any tabs to 4 spaces
    aaa = Left(aaa, tt - 1) + "    " + Mid(aaa, tt + 1)
    GoTo line_1002
'07 november 2002    End If              '05 october 2002
line_1008:
    tt = InStr(aaa, Chr(10)) 'check for line feed january 22 2001
    If tt = 0 Then
        GoTo line_1008d
    End If
    'change any line feeds to nothings
    aaa = Left(aaa, tt - 1) + " " + Mid(aaa, tt + 1)
    ooo = Left(ooo, tt - 1) + " " + Mid(ooo, tt + 1)
    GoTo line_1008
line_1008d:
    ooo = aaa + ""  'save the original chr upper/lower
                    'original input line saved
    save_line = "1008"      'for error handling
'    If Len(aaa) > zzz_len Then
'        zzz_len = Len(aaa)
'        long_line = Left(aaa, 20)   ' maybe fix it in editor?
'    End If
    
    zzz_chrs = zzz_chrs + Len(ooo) + 2  'the cr/lf characters.
    If printed <> "YES" And printed_cnt Mod MAX_CNT * 10000 = 0 Then
        Cls
    End If                          'may 10/00
    If printed <> "YES" And printed_cnt Mod 10000 = 0 Then
        DoEvents        'december 06 2001
        Print "reading "; zzz_cnt
    End If                          'may 10/00
    
    printed_cnt = printed_cnt + 1   'may 10/00    ********************************
    old_line = " " + aaa + " "          '25 March 2003 part of version ver=1.02b
    If uppercase = "Y" Or sscreen_saver = "Y" Then
        aaa = " " + UCase(aaa) + " "
    Else
        aaa = " " + aaa + " "
    End If          '01 october 2002 no need to to put uppercase if all numeric
    
    ooo = " " + ooo + " "   'just to match with aaa
'december 8 2000    If prompt2 = "Q" Then
    If prompt2 = "Q" And uppercase = "N" Then
        aaa = ooo + ""
    End If          'no upper case switch for quicky search
    lll = aaa + ""      'used to show where "P1" matched on in e to exit prompt
    ccc = ooo + ""      'aug 08/99
   
  'no show logic
    
    'might want to have the logic for no-show changed here somewhat
    'ie if anything other that P1 or P then the noshow elements should
    'be cleared (ie count set to 0) etc if "no-show" not the first element
    'this would allow for the replacement of the skip and the complete removal
    'in the text display of anything that shouldn't be shown
    'these notes were done december 30 2000
    'may want the no-show words replaced with "******" so the search will fail
    'todo **vip** at some point in time this is censorship???
'    If prompt2 <> "P1" And prompt2 <> "P" Then GoTo line_1009_a   'december 25 2000
    If extract_yes = "YES" Then GoTo line_1009_a 'january 03 2001
    If encript <> "" Then GoTo line_1009_a  'january 03 2001
'        tt1 = InputBox("doshow logic " + aaa, , , 4400, 4500) 'TESTING ONLY
    If InStr(aaa, "DOSHOW") <> 0 Then GoTo line_1009_a '18 August 2003
    For II = 1 To nocount
'        If zzz_cnt > 80 Then
'            Print "noshow(a)="; "*"; zzz_cnt; "*"; noshow(II); "*"; prompt2; "*"; aaa; nocount
' this one can go on and on so allow for an exit here if bad entry
    If debug_photo Then         '12 october 2002
            tt1 = InputBox("testing photo 6", , , 4400, 4500)  'TESTING ONLY
            If Len(tt1) > 0 Then debug_photo = False    'allow for an out here
    End If
'        End If      'testing only january 27 2001
        If Len(noshow(II)) < 1 Then
            GoTo line_1009
        End If
        If (prompt2 = "P1" Or prompt2 = "P") And InStr(aaa, noshow(II)) <> 0 Then
            GoTo input_1000 'skip if noshow found in line
        End If
        If prompt2 = "SS" And InStr(aaa, noshow(II)) <> 0 Then
            GoTo input_1000 'skip if noshow in screen saver also january 27 2001
        End If
        tt = 1  'january 04 2001
        'the following keeps these words from showing in text display??
line_1008m:
        III = InStr(tt, aaa, noshow(II))   'january 03 2001
        If III = 0 Then GoTo line_1009  'january 03 2001
        tt = III + 1
        JJ = Len(noshow(II))            'january 03 2001
        ooo = Left(ooo, III - 1) + String(JJ, "*") + Mid(ooo, III + JJ) 'january 03 2001
        GoTo line_1008m
'the mid statement without the length works the same as the vax basic right statement vms stuff

line_1009:
    Next II
line_1009_a:       'december 25 2000
  'screensaver logic
'        Print "screensave(?)="; "*"; screensave(1); "*"; aaa; "1"; sscreen_saver; prompt2
'        Print "screensave(?)="; "*"; screensave(screencount); "*"; aaa; screencount; sscreen_saver; prompt2
'        Print sscreen_saver
 '       tt1 = InputBox("testing screen saver logic", , , 4400, 4500)  'TESTING ONLY
 '       If tt1 = "X" Or tt1 = "x" Then
 '           GoTo End_32000
 '       End If              'testing only
    
    'if no screen saver jump around this logic
'28 APRIL 2002    If sscreen_saver <> "Y" Then GoTo line_1009b
    If sscreen_saver <> "Y" Or sscreen_saver_ww = "YES" Then GoTo line_1009b
    
    For II = 1 To screencount
        If Len(screensave(II)) < 1 Then
            GoTo input_1000
        End If
        
        If InStr(aaa, screensave(II)) <> 0 Then
            SSS1 = screensave(II)
' inactive  ss_search = "PHOTO"
            ss_search = screensave(II)
'        Print "inlogic="; "*"; screensave(screencount); "*"; aaa; screencount; sscreen_saver; prompt2
'        tt1 = InputBox("testing screen saver logic", , , 4400, 4500)  'TESTING ONLY
            
            GoTo display_all_2000 'show if screensave found in line
        End If
line_1009a:

    Next II
    GoTo input_1000 'october 24 2000 screen saver fix

line_1009b:
    

    If inin = "A" Or inin = "ALL" Then
        GoTo display_all_2000
    End If
    
    'on flash display to screen
    If SAVE_ttt = "F" Then
        GoTo display_all_2000
    End If

    If SAVE_ttt = "C" And hi_lites = "YES" Then
        GoTo display_all_2000   'aug 08/99
    End If

    If inin = "M" Or inin = "MM" Then
        If Left(aaa, Len(SSS1) + 1) <> " " + SSS1 Then
            GoTo input_1000
        End If
    End If

        
    line_match = "Y"    'november 21 2000
                        'if the input strings match this will remain
                        'otherwise it is reset at the input line

    If InStr(aaa, SSS1) <> 0 Then
        GoTo line_1010
    End If
    line_match = ""     'november 21 2000
    GoTo input_1000

line_1010:
    save_line = "1010"    'for error handling
    If SSS2 = "" Then
        GoTo display_all_2000
    End If
    If InStr(aaa, SSS2) <> 0 Then
        GoTo line_1020
    End If
    line_match = ""     'november 21 2000
    GoTo input_1000
line_1020:
    save_line = "1020"    'for error handling
    If SSS3 = "" Then
        GoTo display_all_2000
    End If
    If InStr(aaa, SSS3) <> 0 Then
'09 june 2002        GoTo display_all_2000
        GoTo line_1030
    End If
    line_match = ""     'november 21 2000
    GoTo input_1000
line_1030:
    If SSS4 = "" Then
        GoTo display_all_2000
    End If
    If InStr(aaa, SSS4) <> 0 Then
        GoTo line_1040
    End If
    line_match = ""     'november 21 2000
    GoTo input_1000
'midway1 thru the program appx'==================================================================================

line_1040:
    If SSS5 = "" Then
        GoTo display_all_2000
    End If
    If InStr(aaa, SSS5) <> 0 Then
        GoTo line_1050
    End If
    line_match = ""     'november 21 2000
    GoTo input_1000
line_1050:
    If SSS6 = "" Then
        GoTo display_all_2000
    End If
    If InStr(aaa, SSS6) <> 0 Then
        GoTo display_all_2000
    End If

'---------------------------------------------
    line_match = ""     'november 21 2000
    GoTo input_1000


'match found now do the hi-light here print print print
display_all_2000:
'       xtemp = InputBox(" 11Nov2012a(match found *****) break_num zzz_cnt prompt2=" + prompt2 + " " + CStr(break_num) + " " + rand_str + " cnt=" + CStr(zzz_cnt) + " " + Left(aaa, 10), , , 4400, 4500) 'TESTING ONLY
'            xtemp = InputBox("testing#4cc Match 29Oct2012 rand=" + CStr(rand) + " break_num=" + CStr(break_num) + " group_match=" + CStr(group_match) + " rand_no=" + CStr(rand_no) + " zzz_cnt=" + CStr(zzz_cnt) + " rand_prog=" + CStr(rand_prog) + " aaa=" + Left(aaa, 10), , , 4400, 4500) 'TESTING ONLY
            If mixx = True And rand <> True And (UCase(Left(aaa, 5)) = " PHOT" Or UCase(Left(aaa, 5)) = " XXX.") Then 'this was for testing and showing print
'                frmproj2.Caption = "30Jun2012x -- aaa =" + aaa + " * " + ss_search + " * " + mpg_file + " * " + Picture_Search + " * " + sscreen_saver_ww + " * " + xxx_found
'                ggg = InputBox(" 30Jun2012x ss_search mpg_file picture_search sscreen_saver_ww xxx_found", , , 4400, 4500)  'TESTING ONLY
'                If ggg = "x" Then GoTo End_32000  '30Jun2012
                GoTo line_2130    ''29Jul2012
'                GoTo Photo_continue_prompt    '29Jul2012 testing this now
            End If
'03Nov2012 the if below
        If Left(UCase(Cmd(77)), 5) <> "CHARA" And mixx And rand <> True Then      '**note** do not override what they have if not noCHARA
'        xtemp = InputBox(" 29Oct2012 input prompt #2 mixx= " + CStr(mixx) + "*", " testing Prompt #2   ", , xx1 - offset1, yy1 - offset2)
            MAX_CNT = 10       'lines per screen cmd(1)
            Font.Size = 48  'font size cmd(2) 48 point size
            BackColor = 0   'black background cmd(3)
'12Feb2017            Def_Fore = 15   'white text cmd(4)
            Cmd(5) = 14     'hi-lite color yellow
'            ForeColor = QBColor(Def_Fore)
            'Cmd(21) = 26
            line_len = 260   'cmd(21)  make line length long.. all formatting done by user
            Cmd(31) = ":"   'the highlite this element
            hilite_this = ":" 'cmd(31)
            Context_lines = 0  'cmd(22)
            rand = 0            '24Jun2012
            Cmd(76) = "LINEPAUSE==0.15"
            Cmd(77) = "CHARACTERPAUSE==.0333"
            Cmd(80) = "PROMPTDETAILS"
                    '22Jun2012
    frmproj2.Caption = program_info + " (cls #4)" '18Dec2013
            If Not mixx Then Set Picture = LoadPicture() '12Feb2017
'            Set Picture = LoadPicture() '18Dec2013
            Set Image1.Picture = LoadPicture(Pict_file)        '20Dec2013
   ' Picture2.Picture = Picture1.Picture                '18dec2013 test
            '18Dec2013        Cls     '12Nov2012
            'frmproj2.AutoRedraw = True
        End If                                      'ie they have their own settings no need to override
 '27Aug2010 reset the file start point to where the last "break" record + 1
'do this only once ie 999999 once a match is found and in random group mode set the file search point to the value in temp1
If break_num < zzz_cnt And rand And Left(Cmd(81), 10) = "RAND_GROUP" Then
            Close #OutFile
            II = DoEvents
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            II = DoEvents       'yield to operating system
            zzz_cnt = 0 'just past the first break record
display_all_2001:
            Line Input #OutFile, aaa
            zzz_cnt = zzz_cnt + 1
            If zzz_cnt < break_num Then GoTo display_all_2001
            break_num = 999999
            rand_no = zzz_cnt       'force new rand start to this break_num line
'11Nov2012            rand_str = "random#" + CStr(rand_no) '18Mar2012
'       xtemp = InputBox("29Oct2012 (match reset start) break_num zzz_cnt " + CStr(break_num) + " " + CStr(zzz_cnt), , , 4400, 4500)  'TESTING ONLY
            If mixx = True Then rand = 0       '29Oct2012 switch this back when break found
            GoTo input_1000a         'When first match found and in random group reset the input file and go look for match sequentially
End If
'       xtemp = InputBox("t27Aug2010 (match continuing) break_num zzz_cnt " + CStr(break_num) + " " + CStr(zzz_cnt), , , 4400, 4500)  'TESTING ONLY
'27Aug2010

'    frmproj2.AutoRedraw = True      '04 January 2005
'       xtemp = InputBox("DOUG 1 TESTING SSS1 SSS2 SSS3 " + SSS1 + " " + SSS2 + " " + SSS3 + " " + SSS4 + " " + SSS5 + " " + SSS6, , , 4400, 4500)  'TESTING ONLY
'
'        tt1 = InputBox("doug testing " + aaa, , , 4400, 4500) 'TESTING ONLY
If Left(UCase(Cmd(70)), 10) = "FOREGROUND" Then
    new_delay_sec = 1           'enough time for the main job to get going again
    GoSub line_30300
    SetFocus        '17 September 2004 background job to front.
                    'this only keeps focus till foreground job runs another file
                    'so the background job that is moved to the foreground should not be any longer
                    'than the other job time wise.... (until something else can be done)
End If          '17 September 2004
    save_line = "2000"    'for error handling
     displayed_cnt = displayed_cnt + 1   'october 13/00
    If displayed_cnt Mod 10 = 0 And ddemo = "YES" And prompt2 = "P1" Then
    'pause a bit and do a couple of beeps just to bug them
    For tt = 1 To 3
        For II = 1 To 3000
            For JJ = 1 To 5000
            Next JJ
        Next II
        Beep
            DoEvents
    Next tt
        tt1 = InputBox("Demo copy only e-mail stonedan@telusplanet.net for full version", , , xx1 - offset1, yy1 - offset2) 'TESTING ONLY
    End If
    If printed_cnt > 20000 And printed <> "YES" Then
        Cls
    End If
    printed = "YES"
    printed_cnt = 1         'may 10/00
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = True      'november 10 2001 autoredraw pair-4
    
    If Context = "no" Then      'june 26/99
        GoTo line_2003
    End If
    temp1 = Len(date_displayed) 'june 26/99 ie "6/26/99"
    'PRINT THE DATE PART AND REMOVE IT FROM HE SEARCH STRING
    '   REPLACING IT WITH THE ORIGINAL SEARCH STRINGS
    'print print print done here
    Print Left(aaa, temp1);
    JJ = Len(aaa)
    aaa = Right(aaa, JJ - temp1) + ""
    ooo = Right(ooo, JJ - temp1) + ""
    SSS1 = SAVE_KEEPS1
    SSS2 = SAVE_KEEPS2
    SSS3 = SAVE_KEEPS3      'june 26/99
    SSS4 = SAVE_KEEPS4      '09 june 2002
    SSS5 = SAVE_KEEPS5
    SSS6 = SAVE_KEEPS6
line_2003:
    save_line = "2003"
    cnt = cnt + 1
    tot_disp = tot_disp + 1     'how many to the screen
  
    'context below deals with the "-" and "=" displays
    'and only applies to date formatted files
    'clean up by moving the context = "yes"
 
 
    tttpos = 1
    KEEPS1 = SSS1 + ""
    KEEPS2 = SSS2 + ""
    KEEPS3 = SSS3 + ""
    KEEPS4 = SSS4 + ""      '09 june 2002
    KEEPS5 = SSS5 + ""
    KEEPS6 = SSS6 + ""
    vvv = aaa
    
    If SSS1 = "A" Or SSS1 = "ALL" Then
        vvv = aaa
    End If

    If SSS1 = "A" Or SSS1 = "ALL" Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS1 <> "" And InStr(aaa, SSS1) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS2 <> "" And InStr(aaa, SSS2) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS3 <> "" And InStr(aaa, SSS3) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS4 <> "" And InStr(aaa, SSS4) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS5 <> "" And InStr(aaa, SSS5) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" And SSS6 <> "" And InStr(aaa, SSS6) = 0 Then
        GoTo line_2200
    End If
    If SAVE_ttt = "F" Then
        hi_lites = "YES"
        vvv = aaa
        GoTo line_2004
    End If
'november 21 2000 the following 3 lines
    If SAVE_ttt = "C" And SSS1 <> "" And hi_lites = "YES" And InStr(aaa, SSS1) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    If SAVE_ttt = "C" And SSS2 <> "" And hi_lites = "YES" And InStr(aaa, SSS2) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    If SAVE_ttt = "C" And SSS3 <> "" And hi_lites = "YES" And InStr(aaa, SSS3) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    '09 JUNE 2002
    If SAVE_ttt = "C" And SSS4 <> "" And hi_lites = "YES" And InStr(aaa, SSS4) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    If SAVE_ttt = "C" And SSS5 <> "" And hi_lites = "YES" And InStr(aaa, SSS5) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    If SAVE_ttt = "C" And SSS6 <> "" And hi_lites = "YES" And InStr(aaa, SSS6) = 0 Then
        mult1 = ""
        mult2 = ""
        mult3 = ""
        mult4 = ""
        mult5 = ""
        mult6 = ""
    End If
    
'    If SAVE_ttt = "C" And SSS1 <> "" And hi_lites = "YES"  And _  january 03a 2001
    If SAVE_ttt = "C" And SSS1 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS1) = 0 Then
'        mult1 = ""          'november 21 2000
'        line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If
'    If SAVE_ttt = "C" And SSS2 <> "" And hi_lites = "YES"  And _  january 03a 2001
    If SAVE_ttt = "C" And SSS2 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS2) = 0 Then
 '       mult2 = ""          'november 21 2000
 '       line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If
'    If SAVE_ttt = "C" And SSS3 <> "" And hi_lites = "YES"  And _  january 03a 2001
    If SAVE_ttt = "C" And SSS3 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS3) = 0 Then
  '      mult3 = ""          'november 21 2000
 '       line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If
' 09 JUNE 2002
    If SAVE_ttt = "C" And SSS4 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS4) = 0 Then
  '      mult3 = ""          'november 21 2000
 '       line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If

    If SAVE_ttt = "C" And SSS5 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS5) = 0 Then
  '      mult3 = ""          'november 21 2000
 '       line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If

    If SAVE_ttt = "C" And SSS6 <> "" And (hi_lites = "YES" Or inin = "A") And _
        InStr(aaa, SSS6) = 0 Then
  '      mult3 = ""          'november 21 2000
 '       line_match = ""     'november 21 2000
        GoTo line_2200  'aug 08/99
    End If

    If SAVE_ttt = "C" And hi_lites <> "YES" Then
        hi_lites = "YES"
 '       line_match = "Y"     'november 21 2000
        GoSub Last_lines_15000  'aug 08/99
        vvv = aaa
        GoTo line_2004
    End If

    

'AT THIS POINT BOLD/HILITING IS GOING TO BE DONE
line_2004:
        Context_cnt = -1 'november 10 2000 this fixes the display
        If p2p2 = "S" And search_str = "CC" Then
            prompt2 = "C"           '26 august 2002
            SAVE_ttt = "C"
            cnt = MAX_CNT       'force it to the end by indicating screen full
            previous_picture(previous_count) = zzz_cnt
'            previous_count = previous_count + 1
'            zzz_cnt = zzz_cnt - Context_lines
'        temptemp = InputBox(" 26 august at bold ", "prompt2= " + prompt2 + p2p2 + " " + CStr(zzz_cnt), , xx1 - offset1, yy1 - offset2)
'un comment the line below when working
            GoTo line_2100
        End If
        GoSub sub_12000        'The hilite display subroutine november 9 2000
line_2100:
    save_line = "2100"
'===============================================================================
    If prompt2 = "C" Then       ' see end of if below "end of if below"
        previous_count = previous_count + 1   'october 19 2000
        If previous_count > 100 Then
            previous_count = 1
        End If
        previous_picture(previous_count) = zzz_cnt
    End If
 
     If Context = "yes" And SAVE_SSS = "=" Then
        If date_displayed <> "" Then
            SSS1 = " "
            SSS = date_displayed
            temp1 = InStr(1, SSS, "/")
            SSS2 = Left(SSS, temp1 - 1) + "/"
            temp1 = InStr(temp1 + 1, SSS, "/")
            SSS3 = "/" + Mid(SSS, temp1 + 1, 2)
        End If
     End If
    
    If Context = "yes" And SAVE_SSS = "-" Then     'june 26/99
        SSS1 = " "
        SSS2 = " "
        SSS3 = date_displayed
    End If
    
    If Context = "no" And SSS1 <> "A" Then
     If KEEPS1 <> "" Then
        SSS1 = KEEPS1
     End If
     If KEEPS2 <> "" Then
        SSS2 = KEEPS2
     End If
     If KEEPS3 <> "" Then
        SSS3 = KEEPS3
        End If
     If KEEPS4 <> "" Then       '09 june 2002
        SSS4 = KEEPS4
        End If
     If KEEPS5 <> "" Then
        SSS5 = KEEPS5
        End If
     If KEEPS6 <> "" Then
        SSS6 = KEEPS6
        End If
    End If
'*******************************************************
' major pause prompt on picture display done here
'*******************************************************
line_2130:
    save_line = "2130"      'november 6 2000
'21 March 2004 maybe do some thing to display the time here too
'    If Picture_Search = "YES" Or mpg_file = "YES" Then
    If Picture_Search = "YES" Then
'        If mpg_file = "YES" Then GoTo line_2130a
            
'            frmproj2.Caption = "lll1=" + lll '08 November 2004
        Line_Search = ""
        GoSub Display_pict_17000
'23 November 2004 do the options== stuff here right after the picture shownmaybe
'================================================================================
'        frmproj2.Caption = " testing=" + UCase(lll) + "="     'testing
        temptemp = UCase(lll)       '23 November 2004
    III = InStr(temptemp, "OPTIONS==")    '23 November 2004
    If III <> 0 Then
        JJ = 0
        temptemp = Right(temptemp, Len(temptemp) - III - 8)
'        frmproj2.Caption = " testing=" + temptemp + "="     'testing

More_ops:
        II = InStr(temptemp, "OPT=")
'        frmproj2.Caption = " testing1=" + temptemp + "=" + CStr(JJ)  'testing
        If II <> 0 Then
            temptemp = Right(temptemp, Len(temptemp) - 4) 'strip off the OPT=
            ddd = InStr(temptemp, "OPT=")
'        frmproj2.Caption = " testing1a=" + temptemp + "=" + CStr(ddd)  'testing
            If ddd <> 0 Then
                control_files(JJ + 1) = Left(temptemp, ddd - 1)
'        frmproj2.Caption = " testing1b=" + temptemp + "=" + CStr(JJ)  'testing
                temptemp = Right(temptemp, Len(temptemp) - ddd + 1)
                JJ = JJ + 1
            Else
                control_files(JJ + 1) = temptemp
                temptemp = ""
                JJ = JJ + 1
            End If
            GoTo More_ops
        End If
Pick_op:
        If JJ <> 0 Then
'have it default to option 1 allways
'display the file here
'        Set Picture = LoadPicture(Pict_file)        'Normal Mode
            
'            frmproj2.Caption = " testing2=" + control_file + "="
            II = InputBox("enter option # 1 thru " + CStr(JJ), , "1", 4400, 4500)
            If II > JJ Or II < 1 Then GoTo Pick_op
        End If
        control_file = control_files(II)
        tempdata = Cmd(73)
'            frmproj2.Caption = " testing=" + tempdata + "="
        DoEvents
        If Left(UCase(tempdata), 10) = "FILESWITCH" Then
            GoSub Control_28000     'change the control file from in a *.txt file
            Close #OutFile
            DoEvents
            OutFile = FreeFile
            DoEvents
            TheFile = Trim(Cmd(46))  'Must be a file name here not a number
'            frmproj2.Caption = " test TheFile=" + TheFile + "="
            Open TheFile For Input As #OutFile
            DoEvents
        End If
'            frmproj2.Caption = " testing1x=" + control_file + "=" '23 November 2004
        GoTo input_1000a
    End If                                      '23 November 2004

'================================================================================
'08 November 2004 put the results to a file ie play list generation or if something interesting played random
        If UCase(Left(Cmd(72), 11)) = "RESULTS.TXT" And UCase(Trim(TheFile)) <> "RESULTS.TXT" Then                   '08 November 2004
'    frmproj2.Caption = " ooo1=" + ooo '08 November 2004 testing Just need "xxx." put in front here.
       ResultFile = FreeFile        '08 November 2004
       Open "RESULTS.TXT" For Append Access Write As #ResultFile   '08 November 2004 do this if RESULTS.TXT IN Cmd(72)
        Print #ResultFile, LTrim(ccc)      '08 November 2004   this is the PHOTO stuff
        Print #ResultFile, Line_Search        '08 November 2004 this is the xxx. stuff
        Close ResultFile                            '08 November 2004
'then do a print / write then a close on that file
    End If                                                      '08 November 2004
        
'           frmproj2.Caption = "hey 7g " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '22Mar2016
'line_2130a:                         '21 March 2004
        '
    
    ' P O S T   P H O T O   D I S P L A Y   C L E A N   U P
    ' show time date info before going on to next picture
    ' maybe something else should be done here in the future
    ' ie quote of the day display etc... random...
    '
    '24 March 2003 do some messing with the time display
    'version ver=1.02b time and date hard coded display for now every 4 show time on the 10"s show day too
    'see similar code at PHOTO_DETAIL calls to Last_lines_15000 and sub_12000
    'doing the time display here after photo and text finished
    'should have a few other options coded here for flexability soon.
    'most of the hard coding should be comming from the control.txt file
    'making it easier to center text what ever.
    '
'    If dsp_cnt Mod 4 = 0 Then       'do every 4 records ver=1.03 make switchable
        special_date = "NO"     '01 April 2003 skip additional display delays if logic below used ver=1.05
        time_displayed = "NO"   '03 April 2003 just indicates if the time below is displayed...
        Def_Fore = Cmd(29)  'Use alt color here for display 26 July 2003
    If dsp_cnt Mod 4 = 0 And Cmd(57) = "SHOWTIME" Then       'do every 4 records
'        replace_data = SSS1      '03 April 2003     save SSS1 so it can be replaced a few lines down
        replace_sss1 = SSS1
        replace_sss2 = SSS2
        replace_sss3 = SSS3
        replace_sss4 = SSS4
        replace_sss5 = SSS5
        replace_sss6 = SSS6     '03 April 2003
        time_displayed = "YES"  '03 April 2003
'        Def_Fore = 12       'make the text red???
'        ForeColor = QBColor(Def_Fore)
        Context_lines = 3               '1 less than the PHOTO_DETAIL display
'        ForeColor = QBColor(12)        'red
'        ForeColor = QBColor(10)         'try lime green instead of red above
'        ForeColor = QBColor(14)         'make this yellow and photo_detail lime green
'        ForeColor = QBColor(9)         'make this bright blue instead
'        ForeColor = QBColor(13)         'make this light purple instead
'        ForeColor = QBColor(14)         'make this yellow instead
'        ForeColor = QBColor(10)         'make this lime green (10) instead
        ForeColor = QBColor(11)         'make this pale blue (11) instead
        Font.Italic = False
        Font.Size = 72                  'making time huge
'        Font.Size = 48                  '20 May 2003 switch to this for laptop
        
'        If sscreen_saver = "Y" Then GoSub line_30000    'do a pause again if screen saver mode


'   P A U S E for screen saver mode here

'            frmproj2.Caption = "hey 7h " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004
'01 September 2004        GoSub line_30000    'do a pause again if screen saver mode
'the above caused it to hang on music thumbs that were being played full
            new_delay_sec = Val(Cmd(60))   '01 September 2004
            GoSub line_30300        '01 September 2004
'            frmproj2.Caption = "hey 7i " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004
'changed setting of aaa to below
'            tt1 = InputBox("testing counts tot_s1=" + CStr(tot_s1), , , 4400, 4500)  'TESTING ONLY
'20 August 2003        Cls
        
        II = 0
        ooo = Now
        II = InStr(1, ooo, " PM")       'trim the seconds out of the time display
        If II <> 0 Then
            ooo = Left(ooo, II - 4) + " PM"
        End If
        
        II = InStr(1, ooo, " AM")       'trim the seconds out of the time display
        If II <> 0 Then
            ooo = Left(ooo, II - 4) + " AM"
        End If
        
        II = 0
        If dsp_cnt Mod 10 <> 0 Then
            II = InStr(2, ooo, " ") 'chop off the day month info ie except if ends in 10 ie 20 30 40 etc
'            tt1 = InputBox("testing time=" + CStr(II) + "=" + ooo, , , 4400, 4500) 'TESTING ONLY
        Else
            ooo = Format(Now, "dddd, mmmm dd, yyyy")    'test another date format

            II = InStr(1, UCase(ooo), ", 20")
        
            If II <> 0 Then
                ooo = Left(ooo, II - 1) 'strip off ", 2003" year info (as with seconds) not needed
                                'logic should work till 2100, then the year will show
            End If
         
            II = InStr(1, ooo, ",")     'remove the comma ie Saturday, March 29
            If II <> 0 Then
                ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - II)
            End If                      '29 march 2003
            
            II = 0
            Font.Size = 72
'            Font.Size = 48          '20 May 2003 switch to this for laptop display
            Context_lines = 3
'...???            SSS1 = Right(ooo, 2)        '30 March 2003      ver=1.04 hilite the day (30) in "Sunday March 30"
            SSS1 = Right(ooo, 2)        '30 March 2003      ver=1.04 hilite the day (30) in "Sunday March 30"
            SSS2 = ""
            SSS3 = ""
            SSS4 = ""
            SSS5 = ""
            SSS6 = ""           '03 April 2003
        End If
        ooo = Right(ooo, Len(ooo) - II)    ' + CStr(dsp_cnt)
        ooo = "  " + ooo        'center the date time display a bit
        If Len(ooo) < 12 Then ooo = "         " + ooo  'ie the time only center 10:52 AM stuff
                            'when too many spaces added above data disappeared ???^
'        dsp_cnt = Len(ooo)
        
'            aaa = "sure would be nice if other stuff prints" + Space(10)   'testing the fix for all characters to print only
        aaa = "dummydummydummydummydummydummy"    'some how forces the full date time to print????
        aaa = UCase(ooo) '+ "-----------------------------------------"              'see if this changes match hi-liting
        
 
 ' testing the hi-liting of the time element..
 '28 March 2003 should be able to hi-lite the following in the time some how
 '       KEEPS1 = "AM"
 '       KEEPS2 = "PM"
        If InStr(1, aaa, " PM") <> 0 Then
'...???            SSS1 = "PM"
            SSS1 = "PM"
            SSS2 = ""
            SSS3 = ""
            SSS4 = ""
            SSS5 = ""
            SSS6 = ""      '03 April 2003
        End If                  '30 March 2003  ver=1.04
        If InStr(1, aaa, " AM") <> 0 Then
'...???            SSS1 = "AM"
            SSS2 = ""
            SSS3 = ""
            SSS4 = ""
            SSS5 = ""
            SSS6 = ""      '03 April 2003
            SSS1 = "AM"
        End If                  '30 March 2003  ver=1.04 done with 2 if statements re the sss1 20 or so lines above
'        SSS1 = " PM"
 '       SSS2 = "PM"             'see if this is hi-lited
 '       s1len = 2
 '       s2len = 2
 '  messing with the above resulted in the hi-liting of the matches disappearing...
 '  in screen saver photo_detail mode.
'            tt1 = InputBox("testing time display keep_sss1=" + keep_SSS1, , , 4400, 4500) 'TESTING ONLY
'            tt1 = InputBox("testing time display ooo=" + ooo, , , 4400, 4500)  'TESTING ONLY
'            tt1 = InputBox("testing time display array_pos=" + CStr(array_pos), , , 4400, 4500)  'TESTING ONLY
        
        
            tempdata = ooo                  '31 March 2003
'03 April 2003            tempdata = " " + ooo                '01 April 2003 ver=1.05 make it shift one to the right before it disappears to the left
        GoSub Last_lines_15000
'21 March 2004        If mpg_file <> "YES" Then GoSub Last_lines_15000
        GoSub sub_12000     'do the bolding and high-liting hi-liting here
            new_delay_sec = Val(Cmd(27))   '03 September 2004
            GoSub line_30300        '03 September 2004
'            tt1 = InputBox("testing=" + CStr(new_delay_sec) + "=", , , 4400, 4500) 'TESTING ONLY 26 November 2004

'30 March 2004 may need a reset of hold_sec here
'        delay_sec = Val(Cmd(27))        '30 March 2004
'        If mpg_file = "YES" Then
'            new_delay_sec = Val(Cmd(27))
'            GoSub line_30300            '30 March 2004
'        End If                          '30 March 2004

'only if screen saver mode
        If dsp_cnt Mod 10 = 0 Then
            special_date = "YES"        '01 April 2003      ver=1.05
'            If sscreen_saver = "Y" Then GoSub line_30000    'do a extra pause again on "weekday Month dd" display
            new_delay_sec = Val(Cmd(27))    '03 September 2004
            If sscreen_saver = "Y" Or sscreen_saver_ww = "YES" Then GoSub line_30300    '03 September 2004
'03 September 2004            If sscreen_saver = "Y" Or sscreen_saver_ww = "YES" Then GoSub line_30000    '03 April 2003
            tempdata = "               " + tempdata     '03 August 2003 start it a little farther over
            f = Len(tempdata)

'                        tt1 = InputBox("testing sscreen_saver=" + sscreen_saver + " " + sscreen_saver_ww, , , 4400, 4500) 'TESTING ONLY
'            For JJ = II To 1 Step -1
            For III = f To 0 Step -1
                
                ooo = Right(tempdata, III)
                aaa = UCase(ooo)
'                        tt1 = InputBox("testing ooo=" + ooo + CStr(f) + CStr(III), , , 4400, 4500) 'TESTING ONLY
'20 August 2003                Cls
                GoSub Last_lines_15000
                GoSub sub_12000
'                delay_sec = 0.1
            new_delay_sec = 0.1    '03 September 2004
            GoSub line_30300            '03 September 2004
'                GoSub line_30000            '31 March 2003    have it move off the screen 1 chr at a time
            Next III
            delay_sec = Val(Cmd(27))        '31 March 2003
        End If
        Font.Size = Val(Cmd(2))     'reset it back from 72 above
    End If  'end of Picture_Search = "YES"
'21 March 2004
'                        tt1 = InputBox("testing ppp=" + ooo + CStr(f) + CStr(III), , , 4400, 4500) '22Mar2016

'03 April 2003        SSS1 = ""               're the am - pm hi-light above  30 march 2003   ver=1.04
    If time_displayed = "YES" Then
'        SSS1 = replace_data            '03 April 2003
        SSS1 = replace_sss1
        SSS2 = replace_sss2
        SSS3 = replace_sss3
        SSS4 = replace_sss4
        SSS5 = replace_sss5
        SSS6 = replace_sss6             '03 April 2003
    End If
    
    If SSS1 = "AM" Or SSS1 = "PM" Then SSS1 = ""        '03 April 2003
    '24 March 2003 ver=1.02b ^

        DoEvents
        If pp_entered = "YES" Then
             GoTo What_50
        End If              'november 6 2000
            
        'skip the pause below if no xxx. found
'            frmproj2.Caption = "testing 21Sep2010 =" + disp_file + "*" + hold_sss1 + "*" + sscreen_saver + "*" + screen_saver_ww + "*" + interrupt_prompt2 '20 November 2004
'03Jul2012 note this section loop for the mixx stuff ----------------------------------------------------
        If InStr(1, UCase(Line_Search), "XXX.") <> 0 Then
'maybe display the last text line if match is photo "P1"
'this has to show for laptops and pc's maybe based on
'lines per inch this must show at 7500 and 7400 for both screens
        tempss = "" 'must have the input string complete
'        If Test1_str = "P1" Or Test1_str = "P" Or Test1_str = "SS" Then     '11Jul2016 add the ss
        If Test1_str = "P1" Or Test1_str = "P" Then
            tempss = lll        'lll is the upper case line
'save the clipboard for later paste
 '           Clipboard.SetText tempss
        End If
'itwasheredoug
        disp_file = Pict_file
line_2150:
        JJ = InStr(1, disp_file, "\")
        If JJ <> 0 Then
            disp_file = Right(disp_file, Len(disp_file) - JJ)
            GoTo line_2150
        End If
        
    disp_file = Right(Line_Search, Len(Line_Search) - 4)     '21Sep2010 maybe this will fix it? Password? url? pic, video make sure they all work
           frmproj2.Caption = "found 21Sep2010 =" + disp_file + "*" + Line_Search + "*" + sscreen_saver + "*" + screen_saver_ww + "*" + interrupt_prompt2 '20 November 2004
'16Aug2016 test the above frm display
    'Screen Saver mode pause takes place below
        
        If sscreen_saver = "Y" Then
            'GoSub line_30000           01 April 2003
'14 July 2003 this is where the pause after the movies was showing the small text (not needed) below
'14 july 2003            If special_date <> "YES" Then GoSub line_30000  '01 April 2003 ver=1.05
            new_delay_sec = Val(Cmd(27))    '03 September 2004
            If line_delay_sec <> 0 Then new_delay_sec = line_delay_sec          '26 November 2004
'            tt1 = InputBox("testing=" + CStr(new_delay_sec) + "=", , , 4400, 4500) 'TESTING ONLY 26 November 2004
'03 September 2004 following line completely removed... when doing music it pauses when it should not
            If special_date <> "YES" And motion_yn <> "YES" Then GoSub line_30300  '01 April 2003 ver=1.05
'03 September 2004            If special_date <> "YES" And mpg_file <> "YES" Then GoSub line_30000  '01 April 2003 ver=1.05
'            Set Picture = LoadPicture()           '03 August 2003 testing (clear em both) lp#2
            Set Image1.Picture = LoadPicture        '03 August 2003 testing
'       xtemp = InputBox(" testing doug#11  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
            GoTo line_2155
        End If
    If screen_capture = "YES" Then
'03 September 2004        delay_sec = 5      'march 15 2001
'03 September 2004        GoSub line_30000
'03mar2017            new_delay_sec = 5    '03 September 2004
            GoSub line_30300            '03 September 2004
    End If
'allow for the keeping of the original file name on the copy
'    copy_photo = "YES"      '11Jul2016 just a hard code test for wmv transfer
    If copy_photo = "YES" Then      'april 08 2001
        temps = Pict_file + ""
next_slasha:
        II = InStr(temps, "\")
        If II = 0 Then GoTo no_slash
        III = III + II
        temps = Mid(temps, II + 1)
        GoTo next_slasha
no_slash:
    old_pict = Mid(Pict_file, III + 1)
    End If

     'march 17 2001
'25Aug2016        xtemp = InputBox(" testing doug 12jul2016", "testing Prompt   " + Cmd(23) + disp_file, xx1 - offset1, yy1 - offset2)
'   copy_photo = "YES"      '11Jul2016
    If copy_photo = "YES" Then
            temps = CStr(photo_cnt + 1)
            If photo_cnt + 1 < 10 Then temps = "0" + temps
           II = InStr(Pict_file, ".")
           temps = photo_dir + photo_file + temps + Mid(Pict_file, II)
'        xtemp = InputBox(" testing doug 12jul2016", " testing Prompt " + temps, , xx1 - offset1, yy1 - offset2)
            If photo_file = "" Then         'april 08 2001
                temps = photo_dir + old_pict
'        xtemp = InputBox(" testing 22 nov " + old_pict, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
            End If
'15Dec2012        xtemp = InputBox(" copy picture to " + temps + " <Y>", "Copy Prompt   ", Pict_file, xx1 - offset1, yy1 - offset2)
        If copy_all = "Y" Then
           xtemp = "Y"
           GoTo line_2150a
        End If              '15Dec2012
        xtemp = InputBox(" (doall) or copy picture to " + temps + " <Y>", "Copy Prompt or (doall)  ", Pict_file, xx1 - offset1, yy1 - offset2)
        If UCase(xtemp) = "DOALL" Then
            copy_all = "Y"
            xtemp = "Y"
        End If              '15Dec2012
'22 November 2006        xtemp = InputBox(" copy picture to " + temps + " <Y>", "Copy Prompt   ", long_pict_file, xx1 - offset1, yy1 - offset2)
'        xtemp = InputBox(" copy picture to " + temps + " <Y>", "Copy Prompt   ", long_pict_file, xx1 - offset1, yy1 - offset2) '12Jul2016 testing
line_2150a:         '15Dec2012
'        Pict_file = Cmd(23) + Right(disp_file, III + 1)     '12jul2016
        
'15Aug2016        If xtemp = "Y" Or xtemp = Pict_file Then
        If xtemp = "Y" Or xtemp = Pict_file Then
            photo_cnt = photo_cnt + 1
            temps = CStr(photo_cnt)
'            temps = Right("00000" + temps, 5) '31Mar2008
            temps = Right("00000" + temps, 5) '11Jul2016
'31Mar2008 If photo_cnt < 10 Then temps = "0" + temps
            temps = photo_dir + photo_file + temps + Mid(Pict_file, II)
            If photo_file = "" Then         'april 08 2001
                temps = photo_dir + old_pict
            End If
'        xtemp = InputBox("testing file =" + temps, , , 4400, 4500) 'TESTING ONLY
line_2153:      '09Aug2016
        ytemp = disp_file + ""          '14Jul2016 needed to update III maybe
'        ytemp = Pict_file + ""  '15Aug2016 needed to update III maybe
more_slasha:
        II = InStr(ytemp, "\")
        If II = 0 Then GoTo nomore_slash
        III = III + II
        ytemp = Mid(ytemp, II + 1)
        GoTo more_slasha
nomore_slash:
'        xtemp = InputBox(" 18jul2016 copy picture " + tempss + " <Y>", "Copy Prompt   ", disp_file, xx1 - offset1, yy1 - offset2) '18Jul2016 testing
On Error GoTo copy_problem  'march 18 2001
            If Not mixx Then Set Picture = LoadPicture()           '12Feb2017
'            Set Picture = LoadPicture() 'march 18 2001 maybe this will do it lp#3
            DoEvents                    'march 18 2001
'12jul2016            FileCopy Pict_file, temps       'copy the file create a copy
'14Jul2016   use the outdir default not cmd(23) below
'test did not work            FileCopy Pict_file, outdir + newname     'testing newname... 14jul2016
'        xtemp = InputBox("testing disp_file 14Jul2015= " + CStr(III) + " " + disp_file, , , 4400, 4500) 'TESTING ONLY 14Jul2016
' 28Aug2016           FileCopy Pict_file, photo_dir + ytemp      'copy the file create a copy changed 18jul2016 copy file
            If photo_dir <> "" Then FileCopy Pict_file, photo_dir + ytemp      'copy the file create a copy changed 18jul2016 copy file
'25Aug2016        xtemp = InputBox("testing old_pict 14Jul2016=" + photo_dir + " *** " + ytemp, , , 4400, 4500) 'TESTING ONLY 14Jul2016
            photo_dir = good_fold      '12Aug2016 restore back to ok mp3 directory
'            FileCopy Pict_file, photo_dir + Right(disp_file, Len(disp_file) - III)      'copy the file create a copy changed 12jul2016
'            FileCopy Pict_file, outdir + Right(disp_file, Len(disp_file) - III + 1)     'copy the file create a copy changed 12jul2016
'        xtemp = InputBox(" copy picture testing " + temps + " <Y>", "Copy Prompt   ", photo_dir + disp_file, xx1 - offset1, yy1 - offset2) '12Jul2016 testing
           
'            DoEvents        '22 November 2006
'        xtemp = InputBox("testing file 22 november 2006=" + temps, , , 4400, 4500) 'TESTING ONLY
'        xtemp = InputBox("testing file 22 november 2006=" + long_pict_file, , , 4400, 4500)  'TESTING ONLY
'            Name temps As long_pict_file   '22 November 2006
            tempss = " was=" + disp_file + " " + Format(Now, "ddddd ttttt")
            If photo_file = "" Then
                tempss = ""
            End If          'april 08 2001
'        xtemp = InputBox(" copy picture testing 1" + tempss + " <Y>", "Copy Prompt   ", photo_dir + disp_file, xx1 - offset1, yy1 - offset2) '12Jul2016 testing
            Print #ExtFile, ccc; " (was=)"; disp_file; " "; Format(Now, "ddddd ttttt")     'march 21 2001 add the was info
'12jul2016            Print #ExtFile, ccc; tempss     'april 08 2001
'12jul2016            Print #ExtFile, "xxx." + temps   'march 18 2001
            Print #ExtFile, "xxx." + photo_dir + ytemp   '18jul2016
'            Print #ExtFile, "xxx." + photo_dir + Right(disp_file, Len(disp_file) - III)   '14jul2016
'            Print #ExtFile, "xxx." + Cmd(23) + Right(disp_file, III + 1)   '12jul2016
            On Error GoTo Errors_31000      'march 18 2001
            GoTo input_1000
copy_problem:       'march 18 2001
        xtemp = InputBox("file copy error=" + CStr(Err.Number) + " " + Err.Description, , , xx1 - offset1, yy1 - offset2) 'march 18 2001
        On Error GoTo Errors_31000    'march 18 2001
'22Mar2016        Resume input_1000       'march 18 2001
        GoTo input_1000             '22Mar2016 somehow this ends up with permission denied?
        End If
        If Len(xtemp) < 3 And UCase(xtemp) <> "N" Then GoTo What_50
        GoTo input_1000
    End If          'march 17 2001
Photo_continue_prompt:              '27 july 2002
    '17 September 2004
    If Left(UCase(Cmd(69)), 8) = "HIT_STOP" And InStr(1, UCase(App.Path + App.EXEName), "BACKGRD") <> 0 Then
        GoTo End_32000 '17 September 2004
    End If          '17 September 2004
'march 19 2001       tt1 = InputBox("E to exit " + tempss + " " + disp_file, "Photo continue Prompt", , xx1, yy1)
'            frmproj2.Caption = "testing pcp =" + disp_file + "*" + hold_sss1 + "*" + sscreen_saver + "*" + screen_saver_ww + "*" + interrupt_prompt2 '20 November 2004
'  21Sep2010          frmproj2.Caption = "testing pcp =" + App.EXEName + "*" + hold_sss1 + "*" + sscreen_saver + "*" + screen_saver_ww + "*" + interrupt_prompt2 '20 November 2004
'        frmproj2.Caption = hold_sss1 + "*" + interrupt_prompt2 + "*" + CStr(zzz_cnt) + "*" + interrupt_prompt2 '20 November 2004
'29 November 2006
'            SendKeys "^c"                   '22Jun2016
'            SendKeys "^c"                   '22Jun2016

        tt1 = ""        '29 November 2006
        If Left(UCase(disp_file), 5) = "HTTP:" Then
            If Mid(disp_file, 6, 1) <> "/" Then
                disp_file = Right(disp_file, Len(disp_file) - 5)
            End If                          '29 January 2007
            'when xxx.http: and no / then take everything after it emails etc 29 January 2007
            Clipboard.SetText disp_file     '29 November 2006
            SendKeys "^c"                   '29 January 2007
            tt1 = disp_file                 '29 November 2006
            frmproj2.Caption = " Paste the URL: " + disp_file + "<<<" '29 November 2006
        End If                          '29 November 2006
'29 November 2006        tt1 = InputBox("P or J for previous E to exit or '.' " + ccc + " " + disp_file, "Photo continue Prompt #pcp" + CStr(zzz_cnt), , xx1 - ppoffset1, yy1 - ppoffset2)
'08 February 2008 add the if text_pause <> true then (below)
If text_pause <> True Then      '08 February 2008
' 21Jun2016 tt1 = InputBox("P or J for previous E to exit or '.' " + ccc + " " + disp_file, "Photo continue Prompt #pcp" + CStr(zzz_cnt), tt1, xx1 - ppoffset1, yy1 - ppoffset2)
 tt1 = InputBox("P or J for previous E to exit or '.' " + ccc + " " + disp_file, "Photo continue Prompt #pcp" + CStr(zzz_cnt), disp_file, xx1 - ppoffset1, yy1 - ppoffset2)

        If tt1 = disp_file Then
            tt1 = ""
        End If              '29 January 2007
Else
'    SetFocus            '08 February 2008
    tt1 = ""            '08 February 2008
    new_delay_sec = Val(Cmd(27))
    GoSub line_30300            '08 February 2008
'    GoTo line_3050             '08 February 2008
End If '08 February 2008
'no    If Cmd(45) = App.EXEName Then Cmd(45) = ""  '07 december 2002 allow interrupt to come here and continue
'      xtemp = InputBox("ww test=" + sscreen_saver_ww + "*" + sscreen_saver + "*" + inin + "*" + tt1, , , xx1, yy1) 'march 18 2001
'    If mpg_file = "YES" Then
'        GoTo input_1000             '23 February 2004
'    End If
'20 November 2004 allow for auto run program to continue if return entered here
    If tt1 = "" And interrupt_prompt2 = "WW" Then  '20 November 2004
        sscreen_saver = "Y"
        prompt2 = "WW"      'this one seems to keep it going where before it stopped. now it gets wrong data
        prompt2 = "SS"      'testing this now
'        Test1_str = "P1"   'or maybe p2 find out later
        tt1 = "WW"
        sscreen_saver_ww = "YES"
'        prompt2 = interrupt_prompt2
'        frmproj2.Caption = SSS1 + SSS2 + "xxx*" + interrupt_prompt2 + CStr(zzz_cnt) '20 November 2004
'        prompt2 = Cmd(47)
'        inin = ""   '20 November 2004 test this
'        GoTo input_1000a
        GoTo input_1000
    End If                                      '20 November 2004
        
        If sscreen_saver_ww = "YES" Then
            sscreen_saver_ww = "NO"     '28 april 2002
            sscreen_saver = "N"         '28 april 2002
            inin = ""
'            GoTo What_50
'
        End If
        If tt1 = "." Then
                ppoffset1 = 4500
                ppoffset2 = 5000
                GoTo Photo_continue_prompt
        End If          '27 july 2002
                        'allow for the photo details to be redisplayed
                        'so they can be view on the screen
                        
'03 August 2003        If img_ctrl = "YES" Then
'        If (img_ctrl = "YES" Or line_fit = "FIT") And line_fit <> "REG" Then
        If img_ctrl = "YES" Or line_fit = "FIT" Then
            Set Image1.Picture = LoadPicture        'february 21 2001
'        Else                                '03 August 2003
            If Not mixx Then Set Picture = LoadPicture()           '12Feb2017
'            Set Picture = LoadPicture() '03 August 2003 lp#4
        End If
'        Set Image1.Picture = LoadPicture        '03 August 2003
  
        If Left(UCase(tt1), 2) = "SS" Then
          If Len(tt1) > 2 Then
            Cmd(26) = " " + Right(tt1, Len(tt1) - 2) + " "
            GoSub line_29200  'set up new screen saver elements
          End If        'may 06 2001
            ttt = "P1"
            If stretch_img <> "NO" Then img_ctrl = "YES"        'march 31 2001
            sscreen_saver = "Y"
            SSS1 = ""
            SSS2 = ""
            SSS3 = ""
            SSS4 = ""   '09 JUNE 2002
            SSS5 = ""
            SSS6 = ""
            GoTo input_1000
        End If                  'february 09 2001
        
        If Len(tt1) > 2 Then
            ttt = tt1
            SSS1 = ttt
            SSS2 = ""
            SSS3 = ""
            SSS4 = ""   '09 JUNE 2002
            SSS5 = ""
            SSS6 = ""
            GoTo line_510
        End If                  'february 09 2001
line_2155:
'        Set Picture = LoadPicture() 'clear any picture
        tt1 = UCase(tt1)
        If tt1 = "PP" And ddemo <> "YES" Then
            Test1_str = "P"
            pp_entered = "YES"
            prompt2 = "P"
            GoTo line_2130
        End If              'november 6 2000
        
'19 December 2004        If tt1 = "P" And ddemo <> "YES" Then
        If (tt1 = "P" Or tt1 = "J") And ddemo <> "YES" Then
            tt1 = "P"       '19 December 2004
            rand = 0        '21 august 2002 keep this as we are backing up to previous picture...
            dsp_cnt = dsp_cnt - 1   'may 09 2001
            pp = previous_count - 1
            If pp < 1 Then
                pp = 100
            End If
            Close #OutFile
            II = DoEvents
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            II = DoEvents       'yield to operating system
            
        tt = 3
        If yyy = "P" Then
            tt = MAX_CNT
            yyy = ""
        End If
        For bbb = 1 To previous_picture(pp) - tt
            Line Input #OutFile, aaa
        Next bbb
        Previous_line = aaa
        Line Input #OutFile, aaa
        zzz_cnt = bbb
        previous_count = pp - 1
        
'testing may 07 2001
'        Print "previous_picture(previous_count)zzz_cnt,previous_count"; TheFile; previous_picture(previous_count); "="; zzz_cnt, previous_count
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
'
        GoTo input_1000
        'read up to the previous picture appx cnt
        End If ' end if for tt1 = "P" And ddemo <> "YES"
'======================================================

'maybe check for a for all here october 27 2000
    If tt1 = "A" Then
        SSS1 = "PHOTO"
        SSS = "PHOTO"
        SAVE_SSS = "PHOTO"
        inin = "A"
        SAVE_KEEPS1 = "PHOTO"
        SAVE_KEEPS2 = ""
        SAVE_KEEPS3 = ""
        SAVE_KEEPS4 = ""
        SAVE_KEEPS5 = ""
        SAVE_KEEPS6 = ""
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY 01 January 2005
        
        GoTo input_1000
'        SAVE_ttt = "S"
    End If
         
        If Len(tt1) = 1 And tt1 <> "P" Then
            Def_Fore = Hold_Fore        '27 July 2003
            GoTo Do_Search_110 'all chrs exit
        End If
            Def_Fore = Hold_Fore        '30Jun2012  testing this new line this seems to fix the color switching problem
        If UCase(tt1) = "E" Then
            GoTo Do_Search_110
        End If
        End If ' end if for InStr(1, UCase(Line_Search), "XXX.") <> 0
'03Jul2012 note this section loop for the mixx stuff ----------------------------------------------------

'20 August 2003        Cls
'        xtemp = InputBox(" testing doug#3  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
        GoTo input_1000
    End If ' end if for Picture_Search = "YES"  see end of it above "end of if below"

'================================================================================

'        temptemp = InputBox(" 26 august step a9 ", "prompt2= " + prompt2 + p2p2 + " " + CStr(zzz_cnt), , 7000, 6000)

    'march 31/00
    
    GoTo line_3010

line_2200:
    Context_cnt = -1        'november 10 2000
    save_line = "2200"  'for error handling
 '    Font.Size = 10
'    BackStyle = 1    MESS WITH THIS A BIT BACKGROUND COLOR ETC
 '   Font.Charset = 2
    
'*********************************************************
    'main print to screen here
    'lots of monkeying around to be done here for
    'log type output extract files etc replace.txt files
'*********************************************************
                ' * * * P R I N T * * *
 'august 10/00 if line longer than 1 chop it
'need to change any tabs to 4 characters
'need to see if there is a valid end of line sequence
'before the 72 chrs and print out that by itself.
'so a bit more to be done before this will work correctly
'need to strip any odd ball characters and check for
'carriage return/linefeed sequences in the line also
    'replace any tabs with 4 spaces right here
line_2210:
    tt = InStr(ooo, Chr(9)) 'check for tabs
    If tt = 0 Then
        GoTo line_2220
    End If
    'change any tabs to 4 spaces
    ooo = Left(ooo, tt - 1) + "    " + Right(ooo, Len(ooo) - tt)
    GoTo line_2210
        
line_2220:
   
 'april 10 2001   If InStr(1, ooo, "append start") <> 0 Then
      
'     If InStr(1, ooo, append_start1) <> 0 Then
      If InStr(1, UCase(ooo), UCase(append_start1)) <> 0 Then
         append_start1 = Mid(ooo, InStr(1, UCase(ooo), UCase(append_start1)), Len(append_start1))
        
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4 '09 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append start"
        SSS1 = append_start1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""   '09 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        aaa = ooo + ""
        GoSub sub_12000
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4 '09 JUNE 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_2300
    End If              'hilite any append lines i put in
'april 10 2001    If InStr(1, ooo, "append end") <> 0 Then
'    If InStr(1, ooo, append_end1) <> 0 Then
      If InStr(1, UCase(ooo), UCase(append_end1)) <> 0 Then
         append_end1 = Mid(ooo, InStr(1, UCase(ooo), UCase(append_end1)), Len(append_end1))
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4 '09 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append end"
        
        SSS1 = append_end1
        SSS2 = ""
        SSS3 = ""
        SSS4 = "" '09 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        aaa = ooo + ""
        GoSub sub_12000
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4 '09 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_2300
    End If              'hilite any append lines i put in
    ' with the ucase stuff below
    If (InStr(1, UCase(ooo), UCase(hilite_this)) <> 0 And _
       hilite_this <> "" And hilite_this <> "     ") Then
        hilite_this = Mid(ooo, InStr(1, UCase(ooo), UCase(hilite_this)), Len(hilite_this))
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4 '09 june 2002
        new5 = SSS5
        new6 = SSS6
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001
        SSS1 = hilite_this
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""   '09 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        aaa = ooo + ""
        GoSub sub_12000
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4 '09 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_2300
    End If              'hilite control element Cmd(31) info only hilites data not on matching line


line_2225:
    'november 10 2000
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4 '09 june 2002
        new5 = SSS5
        new6 = SSS6
        
        SSS1 = ""
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""
        SSS5 = ""
        SSS6 = ""
        GoSub sub_12000
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4
        SSS5 = new5
        SSS6 = new6
        
line_2300:
    'lots need to be done around here in the future.
    
 '   Print Left(ooo, 70)
 '   Print Mid(ooo, 71, 70)
 '   Print Mid(ooo, 141, 70)
 '   Print Mid(ooo, 211, 70)
 '   Print Mid(ooo, 281, 70)
 '   Print Mid(ooo, 351, 70)
 '   Print Mid(ooo, 421, 70)
 '   Print Mid(ooo, 491, 70)
 '   Print Mid(ooo, 561, 70)
 '   Print Mid(ooo, 631, 70)
        
 
line_3010:
        
    save_line = "3010"    'for error handling
    If Picture_Search = "YES" And InStr(UCase(Next_line), "XXX.") <> 0 Then
        Next_line = ""
        aaa = Next_line_save
        GoTo line_1002
            'we read ahead looking at the next line
    End If              'march 31/00
    
    If cnt < MAX_CNT Then
        GoTo input_1000
    End If
    yyy = ""
line_3020:
    save_line = "3020"    'for error handling
    If ttt <> "F" Then
        II = DoEvents       'yield to operating system
    End If      'aug 24/99

' placing the InputBox at position 10,000 x 10,000 puts it
   ' off the screen / thus hiding it...
'10 January 2005    If Not text_pause Then  '06 January 2005  just added the if statement around the print below
    If Not text_pause Or p2p2 = "F" Then  '10 January 2005  just added the if statement around the print below
        ForeColor = QBColor(Val(Cmd(5)))
        Print Format(zzz_cnt, "  #########");
    End If              '06 January 2005
'january 05 2001 display "next match" or "next screen" here
'    Print " do you want to continue y/n <y> 'a' for all 'b' for back, '.' for new search"
    tt1 = "MATCH"                     'january 05 2001
    If inin = "A" Then tt1 = "SCREEN" 'january 05 2001
    Test1_str = ""
    If line_len > Val(Cmd(21)) And extract_yes <> "YES" Then
        line_len = Val(Cmd(21)) 'january 06 2001
        Test1_str = "NOWRAP"
    End If
'24 august 2002
    dblEnd = Timer      'get the end time
'    temptemp = InputBox("01 september 2002 " + p2p2 + CStr(zzz_cnt), "flash", , 2000, 2000)
'what the hey ***vip*** todo why does the following fix need to be used as the
'flash failed to work once the line was put in..... fix it when ever
    If p2p2 <> "F" Then yyy = " elap=" + Format(dblEnd - dblStart, "#####0.000") + " Chrs= " + CStr(zzz_chrs)
'01 september 2002    yyy = " elap=" + Format(dblEnd - dblStart, "###0.000") + " Chrs= " + CStr(zzz_chrs)
'    zzz_chrs = 0        '25 august 2002
'06 January 2005 move the text_pause code from below a few lines to here, no need to see this prompt
' when the screen is going to continue on its own
'10 January 2005    If text_pause Then
    If text_pause And p2p2 <> "F" Then
        frmproj2.Caption = " #4 Pause at line= " + CStr(zzz_cnt) '06 January 2005
'23Feb2017             new_delay_sec = Val(Cmd(27))    '03 September 2004
            GoSub line_30300            '03 September 2004
'03 September 2004        GoSub line_30000
'    frmproj2.Caption = program_info + " (cls #5)" '18Dec2013
            Cls '06 September 2004
        yyy = ""
        GoTo line_3050
    End If              '05 october 2002
    
    Print " prompt #4 Enter for next " + tt1 + " y/n <y> 'a' =all, 'b' =back, '.'" + yyy + Test1_str + " Hits="; tot_s1; " "; tot_s2; " "; tot_s3; " "; tot_s4; " "; tot_s5; " "; tot_s6
    ForeColor = QBColor(Def_Fore)
    
    If show_files_yn Then frmproj2.Caption = program_info + random_info + stretch_info + "  " + show_files '24 december 2002

    If SAVE_ttt = "F" And hi_lites <> "YES" Then

'even with the following code the above " End of Screen" line only showed in about
' 80% of the interrupts. It did noticably slow down the flash display february 20 2001
'        For III = 1 To 1000
'            DoEvents    'february 20 2001
'            II = InStr(ooo, "xxx")  'dummy line did a little bit to enable interrupt display but not much
'        Next III        'to allow for the interrupt to display above end of screen line
        GoTo line_3500
    End If
    If SAVE_ttt = "F" Then      '10 January 2005
        new_delay_sec = Val(Cmd(27))
        GoSub line_30300
        Cls
        GoTo line_3050
    End If              '10 January 2005
'06 January 2005
'    If text_pause Then
'            new_delay_sec = Val(Cmd(27))    '03 September 2004
'            GoSub line_30300            '03 September 2004
'03 September 2004        GoSub line_30000
'            Cls '06 September 2004
'        yyy = ""
'        GoTo line_3050
'    End If              '05 october 2002
'06 January 2005
'       * * * C O N T I N U E   D I S P L A Y   * * *
line_3030:              'november 14 2000
    If page_prompt = "NO" Then
        yyy = ""
        If prompt2 = "C" Then
            mess_cnt = mess_cnt + 1 'january 02 2001
            Print #ExtFile, "------------------- next message "; Format(mess_cnt, "#####0"); " ---------------------"
        End If
        GoTo line_3050
    End If          'december 31 2000

'main end of screen prompt here----------------------------------------
    If screen_capture = "YES" Then
'03 September 2004        delay_sec = 5      'march 15 2001
'03 September 2004        GoSub line_30000
            new_delay_sec = 5    '03 September 2004
            GoSub line_30300            '03 September 2004
    End If
line_3040:
        If search_str = "CC" And p2p2 = "S" Then
            tt1 = "P"
            yyy = "P"
            p2p2 = "C"
            prompt2 = "C"
            SAVE_ttt = "C"
            GoTo skip_input_tt1
        'need to read to previous_count after a close and open
        End If                  '26 august 2002
    yyy = InputBox("Do you want to continue y/n <y>", "Continue Prompt", , 20000, 20000)
    dblStart = Timer      'get the start time 24 august 2002
skip_input_tt1:                 '26 august 2002
'    frmproj2.AutoRedraw = False      '04 January 2005
'    frmproj2.Caption = program_info + " (cls #6)" '18Dec2013
    Cls                             'november 10 needed with the autoredraw
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = False      'november 10 2001 autoredraw pair-4
    yyy = UCase(yyy)
        'COULD DO THINGS WITH THIS PROMPT AS WELL (FUTURE)
'april 22 2001 allow for any HH element to be displayed
line_3045:          'october 21 2001
    If hilite_hh = "Y" And yyy = "HH" And hilite_cnt > 0 Then
        For II = 1 To hilite_cnt
        Pict_file = cript2(II)       ' testing this area
        'pg328 the stretch property - if set to true, the picture loaded into
        '  the image control via the Picture property is stretched (see below)
    If debug_photo Then         '12 october 2002
            tt1 = InputBox("testing photo 6a", , , 4400, 4500)  'TESTING ONLY
    End If
'03 August 2003 do same here Meebee even allow for size to determine FIT or REG ???? another option ????
        
    line_fit = ""                '03 August 2003
    If InStr(1, UCase(aaa), "FIT==") <> 0 Then
            line_fit = "FIT"
    End If                  '03 August 2003
    
    If InStr(1, UCase(aaa), "REG==") <> 0 Then
            line_fit = "REG"
    End If                  '03 August 2003
    
    If (img_ctrl = "YES" Or line_fit = "FIT") And line_fit <> "REG" Then
        Set Image1.Picture = LoadPicture(Pict_file) 'Stretch Mode
'        xtemp = InputBox(" testing doug#2  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

    If (img_ctrl <> "YES" And line_fit <> "FIT") Or line_fit = "REG" Then
'    If img_ctrl <> "YES" Or line_fit = "REG" Then
       Set Picture = LoadPicture(Pict_file)        'Normal Mode
'        xtemp = InputBox(" testing doug#1  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

        
'-------------
'03 August 2003        If img_ctrl = "YES" Then
'            Set Image1.Picture = LoadPicture(Pict_file)
'        Else
'            Set Picture = LoadPicture(Pict_file)
'        End If
        xtemp = InputBox("show picture and wait", "Continue Prompt", , 20000, 20000)
        
'03 august 2003        If img_ctrl = "YES" Then
'            Set Image1.Picture = LoadPicture()
'        Else
'            Set Picture = LoadPicture()
'        End If
    If (img_ctrl = "YES" Or line_fit = "FIT") And line_fit <> "REG" Then
        Set Image1.Picture = LoadPicture() 'Stretch Mode
'        xtemp = InputBox(" testing doug#4  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

    If (img_ctrl <> "YES" And line_fit <> "FIT") Or line_fit = "REG" Then
'    If img_ctrl <> "YES" And line_fit <> "FIT" Then
'    If img_ctrl <> "YES" Or line_fit = "REG" Then
            If Not mixx Then Set Picture = LoadPicture()           '12Feb2017
'        Set Picture = LoadPicture()        'Normal Mode lp#5
'        xtemp = InputBox(" testing doug#5  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

        
        Next II
        hh_cnt = zzz_cnt
        hilite_cnt = 0
        yyy = "B"       'may want to make this "V" todo **vip**
'testing to see if I can get the screen re-displayed(text)
'        GoTo line_3040
    End If                              'april 22 2001
    hilite_cnt = 0                      'april 22 2001

'    If yyy = "A" Then
    If yyy = "A" Or (yyy = "" And inin = "A") Then 'january 04 2001
        inin = "A"
        SSS1 = KEEPS1
        SSS2 = KEEPS2
        SSS3 = KEEPS3
        SSS4 = KEEPS4
        SSS5 = KEEPS5
        SSS6 = KEEPS6
        
        hi_lites = "YES"
'        hi_lites = ""
        yyy = ""
        prompt2 = "C"
        SAVE_ttt = "C"
        line_match = ""    'this gets rid of the first one printing hilited
                            ' if the previous page had one hilited???
    End If                  'january 03a 2001 testing
    If inin <> "A" Then
        array_pos = 0
        array_prt = 0
    End If                  'january 21 2001
line_3050:
 '       Print "tt1="; tt1; "=previous_count ="; previous_count
 '       tt1 = InputBox("testing only", , , 4400, 4500)  'TESTING ONLY
 '       If tt1 = "x" Or tt1 = "X" Then
 '           GoTo End_32000
 '       End If          'testing
        tt1 = yyy
        
        If tt1 = "F" Then
            SAVE_ttt = "F"      'november 14 2000
            yyy = "."           'december 25 2000
            tt1 = "."           'december 25 2000
'            GoTo line_3030     'december 25 2000
        End If
'19 December 2004        If tt1 = "P" And ddemo <> "YES" Then
        If (tt1 = "P" Or tt1 = "J") And ddemo <> "YES" Then
            tt1 = "P"           '19 December 2004
'--------------------------------------------------------------
'november 8 2000
            If prompt2 = "Q" Then
                prompt2 = "C"
                SAVE_ttt = "C"
                SSS1 = UCase(SSS1)
                SSS2 = UCase(SSS2)
                SSS3 = UCase(SSS3)
                SSS4 = UCase(SSS4)
                SSS5 = UCase(SSS5)
                SSS6 = UCase(SSS6)
            End If
'--------------------------------------------------------------
'26 august 2002            pp = previous_count - 1
'   this logic only does the "CC" search once could be reset for each new search
'   but once the search is half way thru the file there may be no time savings
'   at least on the first search if no match it will be much much quicker
'   26 august 2002
'        temptemp = InputBox(" 26 august at previous ", "prompt2= " + prompt2 + p2p2 + " " + CStr(zzz_cnt), , xx1 - offset1, yy1 - offset2)
            pp = previous_count     '26 august 2002
            If search_str <> "CC" Then pp = previous_count - 1
            'same logic for previous match as previous picture
line_3200:
            If pp < 1 Then
                pp = 100
            End If
            Close #OutFile
            II = DoEvents
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            II = DoEvents       'yield to operating system
'if previous_picture count is on the same page go 1 more back
        If previous_picture(pp) + MAX_CNT > zzz_cnt And previous_picture(pp - 1) <> 0 Then
            pp = pp - 1
            GoTo line_3200
        End If
        
        For bbb = 1 To previous_picture(pp) - MAX_CNT
            Line Input #OutFile, aaa
        Next bbb
        Previous_line = aaa
        Line Input #OutFile, aaa
        zzz_cnt = bbb
        previous_count = pp - 1
'        hi_lites = "NO"     'the only difference between text and photo
'19 August 2003        Cls
        yyy = ""
'        Print "previous_picture(previous_count)zzz_cnt,previous_count"; TheFile; previous_picture(previous_count); "="; zzz_cnt, previous_count
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
'
'        GoTo input_1000
        'read up to the previous picture appx cnt
        End If
    hi_lites = "NO"     'flash display
    printed_cnt = 1         'may 10/00
    printed = "NO"
    
    If SAVE_ttt = "C" Then
    For temp1 = 1 To MAX_CNT
        Context_text(temp1) = ""     'aug 08/99
    Next temp1
    End If                      'aug 08/99
    
    'allow for full display if "A" entered at this prompt
    If yyy = "A" Then
        SSS1 = "A"
        SSS = "A"
        SAVE_SSS = "A"
        inin = "A"
        SAVE_KEEPS1 = "A"
        SAVE_KEEPS2 = ""
        SAVE_KEEPS3 = ""
        SAVE_KEEPS4 = ""
        SAVE_KEEPS5 = ""
        SAVE_KEEPS6 = ""
        
        yyy = ""
'        SAVE_ttt = "S"
        'december 22 2000
        prompt2 = "C"   'december 22 2000
        SAVE_ttt = "C"  'december 22 2000
    End If
'allow them to back up 1 page or what ever number of lines in Cmd(23)
line_3300:
'january 06 2001    If yyy = "B" Then
    If yyy = "B" Or yyy = "V" Then
'        Print "previous_picture(previous_count)zzz_cnt,previous_count"; TheFile; previous_picture(previous_count); "="; zzz_cnt, previous_count
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
        If yyy = "V" Then
            line_len = 5000   'january 06 2001
            yyy = "B"
        End If
        If prompt2 = "Q" Then
            prompt2 = "C" 'october 22 2000
            If previous_picture(previous_count) <> 0 Then
                zzz_cnt = previous_picture(previous_count)
            End If
        End If
'--------------------------------------------------------
'january 04b 2001 to allow for the B to keep the hilites
        inin = "A"
        SSS1 = KEEPS1
        SSS2 = KEEPS2
        SSS3 = KEEPS3
        SSS4 = KEEPS4
        SSS5 = KEEPS5
        SSS6 = KEEPS6
        hi_lites = "YES"
'        hi_lites = ""
        yyy = ""
        prompt2 = "C"
        SAVE_ttt = "C"
        line_match = ""    'this gets rid of the first one printing hilited

'--------------------------------------------------------
'january 04b 2001 to allow for the B to keep the hilites
'        SSS1 = "A"
'        SSS = "A"
'        SAVE_SSS = "A"
'        inin = "A"
'        SAVE_KEEPS1 = "A"
'        SAVE_KEEPS2 = ""
'        SAVE_KEEPS3 = ""
'        yyy = ""
'--------------------------------------------------------
        Close #OutFile
        DoEvents
        Open TheFile For Input As #OutFile
        II = DoEvents       'yield to operating system
        If cnt > MAX_CNT Then
            tt = cnt - MAX_CNT
            wrap_cnt = wrap_cnt - tt
            cnt = MAX_CNT 'january 08a 2001
        End If
  '     xtemp = InputBox("testing prompt_=" + CStr(zzz_cnt) + " " + CStr(wrap_cnt) + " " + CStr(cnt), "test", , 4400, 4500) '
 '   xtemp = InputBox("testing prompt=" + AllSearch(1) + "*" + ttt + " " + SSS1 + " " + SSS2 + " " + SSS3 + " " + CStr(zzz_cnt), "test", , 4400, 4500) '
'     If UCase(xtemp) = "X" Then GoTo End_32000
    
'        For bbb = 1 To zzz_cnt - MAX_CNT - MAX_CNT + wrap_cnt - 1
        If hilite_hh = "Y" Then zzz_cnt = hh_cnt + MAX_CNT     'april 22 2001
        For bbb = 1 To zzz_cnt - MAX_CNT - MAX_CNT + wrap_cnt - MAX_CNT + cnt
            Line Input #OutFile, aaa
        Next bbb
        zzz_cnt = bbb - 1
    End If
    wrap_cnt = 0            'january 07 2001
    
    If yyy = "." Then
        TheSearch = "."
        ttt = "."
        GoSub Search_26000
        If prompt2 = "C" Then
            SSS = UCase(ttt)
        Else
            SSS = ttt
        End If          'december 29 2000
        If Len(SSS) < 2 Then GoTo What_50   'january 05 2001
        SSS2 = ""       'december 28 2000
        SSS3 = ""       'december 28 2000
        SSS4 = ""       '09 june 2002
        SSS5 = ""
        SSS6 = ""
'december 22 2000
    If SSS = "A" Then
        SSS1 = "A"
        SSS = "A"
        SAVE_SSS = "A"
        inin = "A"
        SAVE_KEEPS1 = "A"
        SAVE_KEEPS2 = ""
        SAVE_KEEPS3 = ""
        SAVE_KEEPS4 = ""
        SAVE_KEEPS5 = ""
        SAVE_KEEPS6 = ""
        yyy = ""
'        SAVE_ttt = "S"
        'december 22 2000
        prompt2 = "C"   'december 22 2000
        SAVE_ttt = "C"  'december 22 2000
    End If
        
        'do the parse again
    'PARSE THE SSS STRING INTO PARTS SSS1 SSS2 AND SSS3
line_3350:
    i = InStr(SSS, "  ")
    If i <> 0 Then
        SSS = Left(SSS, i) + Mid(SSS, i + 2)
        GoTo line_3350
    End If
    
    i = InStr(SSS, sep)
    SSS1 = SSS + ""
    s1len = Len(SSS1)
    inin = SSS + ""         'june 13/99
    If i = 0 Then
        GoTo line_3400
    End If
    SSS1 = Left(SSS, i - 1)
    s1len = Len(SSS1)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

    i = InStr(SSS, sep)
    SSS2 = SSS
    s2len = Len(SSS2)
    If i = 0 Then
        GoTo line_3400
    End If
    SSS2 = Left(SSS, i - 1)
    s2len = Len(SSS2)
    j = Len(SSS)
    SSS = Right(SSS, j - i)
 '=========================================
    i = InStr(SSS, sep)
    SSS3 = SSS
    s3len = Len(SSS3)
    If i = 0 Then
        GoTo line_3400
    End If
    SSS3 = Left(SSS, i - 1)
    s3len = Len(SSS3)
    j = Len(SSS)
    SSS = Right(SSS, j - i)
 '-----------------------------------------
     i = InStr(SSS, sep)
    SSS4 = SSS
    s4len = Len(SSS4)
    If i = 0 Then
        GoTo line_3400
    End If
    SSS4 = Left(SSS, i - 1)
    s4len = Len(SSS4)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

 '-----------------------------------------
     i = InStr(SSS, sep)
    SSS5 = SSS
    s5len = Len(SSS5)
    If i = 0 Then
        GoTo line_3400
    End If
    SSS5 = Left(SSS, i - 1)
    s5len = Len(SSS5)
    j = Len(SSS)
    SSS = Right(SSS, j - i)

'------------------------------------------
    SSS6 = Right(SSS, j - i)
    s6len = Len(SSS6)


line_3400:
    GoSub line_14500        'check for imbedded spaces in search strings
                            'january 19 2001

        yyy = ""

    End If          'end if for "." search string recall

    'allow for a stop to see if screen remains
    If yyy = "S" Then
        Close #OutFile
        II = DoEvents
        GoTo End_32000
    End If
line_3500:
    save_line = "3500"
    
'17 September 2004 dougheredoughere
    If Left(UCase(Cmd(69)), 8) = "HIT_STOP" And InStr(1, UCase(App.Path + App.EXEName), "BACKGRD") <> 0 Then
        GoTo End_32000 '17 September 2004
    End If          '17 September 2004
line_4000:
    save_line = "4000"
    cnt = 0
    If yyy <> "Y" And yyy <> "" Then

        Close #OutFile
        Close #ExtFile          'november 14 2000
        line_len = Val(Cmd(21)) 'november 14 2000
        extract_yes = "NO"      'november 14 2000
        II = DoEvents       'yield to operating system
        GoTo Do_Search_110
 '       GoTo File_40      'january 01 2001
    End If
'20 August 2003 the cls was moved to the last_lines_1500 routine
'20 August 2003    Cls     'clear the current form
    If SAVE_ttt = "F" Then Cls     '20 august 2003a

    GoTo input_1000
    
    
'Subroutines start here ----------------------------
sub_10000:
   sslen = 0
   s1 = 0
   s2 = 0
   s3 = 0
   s4 = 0       '23 june 2002
   s5 = 0
   s6 = 0
   ss = 1000

    If SSS1 <> "" Then
       s1 = InStr(aaa, SSS1)
    End If
    If SSS2 <> "" Then
        s2 = InStr(aaa, SSS2)
    End If
    If SSS3 <> "" Then
        s3 = InStr(aaa, SSS3)
    End If
    If SSS4 <> "" Then
        s4 = InStr(aaa, SSS4)   '23 june 2002
    End If
    If SSS5 <> "" Then
        s5 = InStr(aaa, SSS5)
    End If
    If SSS6 <> "" Then
        s6 = InStr(aaa, SSS6)
    End If

   If s1 < ss And s1 <> 0 Then
       ss = s1
       sslen = s1len
   End If
   If s2 < ss And s2 <> 0 Then
       ss = s2
       sslen = s2len
   End If
   If s3 < ss And s3 <> 0 Then
       ss = s3
       sslen = s3len
   End If
   If s4 < ss And s4 <> 0 Then
       ss = s4              '23 june 2002
       sslen = s4len
   End If
   If s5 < ss And s5 <> 0 Then
       ss = s5
       sslen = s5len
   End If
   If s6 < ss And s6 <> 0 Then
       ss = s6
       sslen = s6len
   End If

'    If ss = s1 Then
'december 5 2000        SSS1 = ""
'    End If
'    If ss = s2 Then
'december 5 2000        SSS2 = ""
'    End If
'    If ss = s3 Then
'december 5 2000        SSS3 = ""
'    End If

    If ss = 1000 Then
            ss = 0
    End If
   
Return

sub_12000:         'hi-lite print subroutine here
'need the tabs to spaces where it is. Spaces may be in the search for matches?
' test to eliminate excess prints
    '20 May 2003 note I may want to de-activate the following change along with
    '            the other one for 20 May 2003 this data is now displayed in
    '            the detail display / caption option....
'    If sscreen_saver = "Y" Then GoTo line_13999      '20 May 2003
'        testprompt = InputBox("testing prompt 12050=" + sscreen_saver + " " + disp_file + " " + aaa, "test", , 4400, 4500)  '
line_12050:     'january 15 2001
    If array_pos <> 0 Then
        data_aaa = aaa + ""
        data_ooo = ooo + ""
        endstuff = "YES"
        array_prt = array_prt + 1
        aaa = array_aaa(array_prt)
        ooo = array_ooo(array_prt)
        SSS1 = KEEPS1       'january 24 2001
        If SSS1 = "A" Then SSS1 = ""    'january 25 2001
        SSS2 = KEEPS2
        SSS3 = KEEPS3       'january 24 2001
        SSS4 = KEEPS4       '09 june 2002
        SSS5 = KEEPS5
        SSS6 = KEEPS6
'        testprompt = InputBox("testing prompt 12050=" + CStr(array_prt) + " " + data_ooo + " " + CStr(array_pos) + " " + array_ooo(array_prt), "test", , 4400, 4500) '
        GoTo line_12120       'january 21 2001
    End If
line_12053:                     'january 21 2001
    endstuff = "NO"             'january 21 2001
    data_aaa = aaa + " "
    data_ooo = ooo + " "
    tot_print = 0
line_12055:
    'below reduce multiple trailing spaces to 1
    JJ = Len(data_aaa)
    If JJ < 2 Then GoTo line_12070
    If Right(data_aaa, 2) = "  " Then
        data_aaa = Left(data_aaa, JJ - 1)
        data_ooo = Left(data_ooo, JJ - 1)
        GoTo line_12055
    End If

line_12070:
    aaa = data_aaa + ""
    ooo = data_ooo + ""
line_12100:
    
    'now do a search for the match strings.
    match_flag = "YES"
'is there a match of the 3 elements in this string YES or NO
    If SSS1 = "" Then match_flag = "NO"
    If SSS1 <> "" And InStr(aaa, SSS1) = 0 Then
        match_flag = "NO"
        GoTo line_12110
    End If
    If SSS2 <> "" And InStr(aaa, SSS2) = 0 Then
        match_flag = "NO"
        GoTo line_12110
    End If
'23 june 2002 add 4 5 and 6 below
    If SSS3 <> "" And InStr(aaa, SSS3) = 0 Then
        match_flag = "NO"
        GoTo line_12110
    End If
    If SSS4 <> "" And InStr(aaa, SSS4) = 0 Then
        match_flag = "NO"
        GoTo line_12110
    End If
    If SSS5 <> "" And InStr(aaa, SSS5) = 0 Then
        match_flag = "NO"
        GoTo line_12110
    End If
    If SSS6 <> "" And InStr(aaa, SSS6) = 0 Then match_flag = "NO"
line_12110:
    If Len(aaa) > line_len + over_lap Then
        GoTo line_13000     'do the line wrap inserts etc
    End If
line_12120:     'return from the wrap logic here if required.
 '   If zzz_cnt > 43180 Then
'        testprompt = InputBox("testing prompt 12120=" + match_flag + "*" + SSS1 + "*" + SSS2 + "*" + SSS3 + "*" + SSS4 + "*" + SSS5 + "*" + SSS6, , , 4400, 4500) 'TESTING ONLY
'         If tt1 = "X" Or tt1 = "x" Then
'            GoTo End_32000
'        End If              'testing only
 '   End If
'29Oct2012 not this one   match_flag            If group_match = True Then xtemp = InputBox("testing#4c1 Match 29Oct2012 break_num=" + CStr(break_num) + " group_match=" + CStr(group_match) + " rand_no=" + CStr(rand_no) + " zzz_cnt=" + CStr(zzz_cnt) + " rand_prog=" + CStr(rand_prog) + " >" + Left(aaa, 8), , , 4400, 4500) 'TESTING ONLY
    If match_flag = "YES" Then GoTo line_12500   '
line_12125:
    
    If Left(UCase(Cmd(77)), 16) = "CHARACTERPAUSE==" Then     '03 January 2005
            new_delay_sec = Val(Right(Cmd(77), Len(Cmd(77)) - 16))
            tempss = ooo                '03 January 2005
            If Left(tempss, 2) = "  " Then tempss = Right(tempss, Len(tempss) - 2) '31Jul2011 so on line wraps the first 2 characters space are removed
            If Left(tempss, 1) = " " Then tempss = Right(tempss, Len(tempss) - 1) '31Jul2011 so on line wraps the first character space is removed
            For temp1 = 1 To Len(tempss)
             Print Left(tempss, 1);
             tempss = Right(tempss, Len(tempss) - 1)
            GoSub line_30300            '03 January 2005 slowed print character pause
            Next temp1
        
    Else
        Print ooo;
    End If                  '03 January 2005
    tot_print = tot_print + Len(ooo)
        If extract_yes = "YES" Then
            Print #ExtFile, ooo;
        End If
 '24 December 2004 put the timer in here???
    If Left(Cmd(76), 11) = "LINEPAUSE==" Then
            new_delay_sec = Val(Right(Cmd(76), Len(Cmd(76)) - 11))
            GoSub line_30300            '24 December 2004
    End If                      '24 December 2004

line_12130:
    If array_prt > 1 Then
        cnt = cnt + 1
        wrap_cnt = wrap_cnt + 1
    End If
             posstring = ""          'december 6 2000
             If showasc = "Y" Then
                 II = Len(ooo)
                 If II > 10 Then II = 10
                 ytemp = Right(ooo, II)
                 xtemp = ""
                 For III = 1 To II
                     xtemp = xtemp + CStr(Asc(Mid(ytemp, III, 1))) + " "
                 Next III
                 posstring = xtemp + "(" + ytemp + ")"
             End If              'december 11 2000
            If showpos = "Y" Then posstring = "*=" + CStr(tot_print) + " " + CStr(cnt)
    Print ; posstring
 '24 December 2004 put the timer in here???
    If Left(Cmd(76), 11) = "LINEPAUSE==" Then
            new_delay_sec = Val(Right(Cmd(76), Len(Cmd(76)) - 11))
            GoSub line_30300            '24 December 2004
    End If                      '24 December 2004
    tot_print = 0
        If extract_yes = "YES" Then
            Print #ExtFile,
        End If
    'check for screen full here 'to be done yet
'at this point need to check for screen full and exit out if need be but do later
line_12135:                 'january 21 2001
    If cnt >= MAX_CNT Then
        If array_prt = array_pos Then
            array_prt = 0
            array_pos = 0
        End If
        GoTo line_14222
    End If
line_12140:
    If array_pos < 1 Then GoTo line_14222
line_12150:
    array_prt = array_prt + 1
line_12160:
    If array_prt > array_pos Then
        array_prt = 0
        array_pos = 0
        If endstuff = "YES" Then
            aaa = data_aaa + ""
            ooo = data_ooo + ""
            GoTo line_12053    'january 21 2001
        End If
        GoTo line_14222
    End If
line_12170:
    aaa = array_aaa(array_prt) + ""
    ooo = array_ooo(array_prt) + ""
    GoTo line_12120
line_12500:
    s1len = Len(SSS1)
    s2len = Len(SSS2)
    s3len = Len(SSS3)
    s4len = Len(SSS4)   '23 june 2002
    s5len = Len(SSS5)
    s6len = Len(SSS6)

    GoSub sub_10000 'check for match in string aaa
    If ss = 0 Then GoTo line_12125
line_12505:
    If ss > 1 Then
'29Oct2012
'            If group_match = True Then xtemp = InputBox("testing#4c2 Match 29Oct2012 match_flag=" + match_flag + " break_num=" + CStr(break_num) + " group_match=" + CStr(group_match) + " rand_no=" + CStr(rand_no) + " zzz_cnt=" + CStr(zzz_cnt) + " rand_prog=" + CStr(rand_prog) + " >" + Left(aaa, 8), , , 4400, 4500) 'TESTING ONLY
    If Left(UCase(Cmd(77)), 16) = "CHARACTERPAUSE==" Then     '03 January 2005
            new_delay_sec = Val(Right(Cmd(77), Len(Cmd(77)) - 16))
            tempss = Left(ooo, ss - 1)               '03 January 2005
            If Left(tempss, 1) = " " Then tempss = Right(tempss, Len(tempss) - 1) '31Jul2011 so on line wraps the first character space is removed
            For temp1 = 1 To Len(tempss)
             Print Left(tempss, 1);
             tempss = Right(tempss, Len(tempss) - 1)
            GoSub line_30300            '03 January 2005 slowed print character pause
            Next temp1
        
    Else
'        Print ooo;
        Print Left(ooo, ss - 1);
             tot_print = tot_print + ss - 1
        If extract_yes = "YES" Then
            Print #ExtFile, Left(ooo, ss - 1);
        End If
    End If                  '03 January 2005
    End If
line_12510:
    II = 0
    If Mid(aaa, ss + sslen - 1, 1) <> " " Then GoTo line_12600

    'if last chr in match a space check if that is the start of another match? below
    If SSS1 = "" Then GoTo line_12600
    If Mid(aaa, ss + sslen - 1, s1len) = SSS1 Then GoTo line_12550
    If SSS2 = "" Then GoTo line_12600
    If Mid(aaa, ss + sslen - 1, s2len) = SSS2 Then GoTo line_12550
    If SSS3 = "" Then GoTo line_12600
    If Mid(aaa, ss + sslen - 1, s3len) = SSS3 Then GoTo line_12550
    If SSS4 = "" Then GoTo line_12600   '23 june 2002
    If Mid(aaa, ss + sslen - 1, s4len) = SSS4 Then GoTo line_12550
    If SSS5 = "" Then GoTo line_12600
    If Mid(aaa, ss + sslen - 1, s5len) = SSS5 Then GoTo line_12550
    If SSS6 = "" Then GoTo line_12600
    If Mid(aaa, ss + sslen - 1, s6len) = SSS6 Then GoTo line_12550
    GoTo line_12600
line_12550:
    'only going to hi-lite up to the next space by reducing count by 1 below
'    If zzz_cnt > 658 Then
'        tt1 = InputBox("testing prompt 12550" + CStr(ss + sslen - 1) + "*" + Mid(aaa, ss + sslen - 1, 4) + "*" + SSS1, , , 4400, 4500) 'TESTING ONLY
'        If tt1 = "X" Or tt1 = "x" Then
'            GoTo End_32000
'        End If              'testing only
'    End If
    sslen = sslen - 1
    II = 1      'so the match counts below will jive
line_12600:
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS1) Then tot_s1 = tot_s1 + 1 'january 01 2001
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS2) Then tot_s2 = tot_s2 + 1 'january 01 2001
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS3) Then tot_s3 = tot_s3 + 1 'january 01 2001
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS4) Then tot_s4 = tot_s4 + 1 '23 june 2002
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS5) Then tot_s5 = tot_s5 + 1 '23 june 2002
   If UCase(Mid(ooo, ss, sslen + II)) = UCase(SSS6) Then tot_s6 = tot_s6 + 1 '23 june 2002
    
    ForeColor = QBColor(Set_Fore)     'set color
    Font.Bold = True
'03mar2017    Font.Underline = True

'26 March 2003 try the below (seemed to work)    Print Mid(ooo, ss, sslen);
'    If Cmd(56) <> "PHOTO_DETAIL" Then Print Mid(ooo, ss, sslen);
'    needed the original display for "C" context hi-liting.... put back
    If Left(UCase(Cmd(77)), 16) = "CHARACTERPAUSE==" Then     '03 January 2005
            new_delay_sec = Val(Right(Cmd(77), Len(Cmd(77)) - 16))
            tempss = Mid(ooo, ss, sslen)              '03 January 2005
            If Left(tempss, 1) = " " Then tempss = Right(tempss, Len(tempss) - 1) '31Jul2011 so on line wraps the first character space is removed
'            If Left(tempss, 1) = " " Then tempss = Right(tempss, Len(tempss) - 1) '31Jul2011 so on line wraps the first character space is removed
            For temp1 = 1 To Len(tempss)
             Print Left(tempss, 1);
             tempss = Right(tempss, Len(tempss) - 1)
            GoSub line_30300            '03 January 2005
            Next temp1
        
    Else
'        Print ooo;
    Print Mid(ooo, ss, sslen);
    End If                  '03 January 2005
    
    tot_print = tot_print + sslen
        If extract_yes = "YES" Then
            Print #ExtFile, Mid(ooo, ss, sslen);
         End If
'    If sscreen_saver = "Y" Then
'26 March 2003 deactivate the following if  part of version ver=1.02b fix required
    If sscreen_saver = "Y" And sscreen_saver = "N" Then
        Font.Size = 24
        ForeColor = QBColor(AltColor) 'make a different color here and below screen saver
'        dsp_cnt = dsp_cnt + 1       'may 09 2001
'26 March 2003 dump the display count for now (skip the whole routine)
'        with the dsp_cnt print and all what the hey anyway
'        Print Mid(ooo, ss, sslen); " "; CStr(dsp_cnt); " "; 'may 09 2001
        Print Mid(ooo, ss, sslen); " "; 'may 09 2001
        Font.Size = Val(Cmd(2))
    End If      'october 24 2000
    ForeColor = QBColor(Def_Fore)    'default color
    Font.Bold = False
    Font.Underline = False
    
    aaa = Mid(aaa, ss + sslen)
    ooo = Mid(ooo, ss + sslen)

'midway2 thru the program appx'==================================================================================

line_12610:
    If Len(aaa) > 0 Then GoTo line_12500
    GoTo line_12130     'no more data

    
    GoTo line_14222

line_13000:          'january 18 2001
    'below get rid of all trailing spaces
    II = Len(aaa)
    If Right(aaa, 1) = " " Then
        aaa = Left(aaa, II - 1)
        ooo = Left(ooo, II - 1)
        GoTo line_13000
    End If
line_13010:         'january 19 2001 see imbedded spaces in search string and replace them in xxx string
    xxx = aaa + ""
'    If zzz_cnt > 20929 Then
'        tt1 = InputBox("testing prompt 13000c" + CStr(Len(aaa)) + " " + SSS1 + " " + match_flag, , , 4400, 4500) 'TESTING ONLY
'        If tt1 = "X" Or tt1 = "x" Then
'            GoTo End_32000
'        End If              'testing only
'    End If
    If match_flag <> "YES" Then GoTo line_13070 'only required if match in line
    If imbedded = "NO" Then GoTo line_13070 'only required if one has a imbedded space
    If s1_imbed = "" Then GoTo line_13020
    II = 1
line_13012:
    JJ = InStr(II, aaa, SSS1)
    If JJ = 0 Then GoTo line_13020
    xxx = Left(aaa, JJ - 1) + s1_imbed + Mid(aaa, JJ + Len(SSS1) - 1)
    II = JJ + 1
    GoTo line_13012
line_13020:
    If s2_imbed = "" Then GoTo line_13030
    II = 1
line_13022:
    JJ = InStr(II, aaa, SSS2)
    If JJ = 0 Then GoTo line_13030
    xxx = Left(aaa, JJ - 1) + s2_imbed + Mid(aaa, JJ + Len(SSS2) - 1)
    II = JJ + 1
    GoTo line_13022
line_13030:
    If s3_imbed = "" Then GoTo line_13040
    II = 1
line_13032:
    JJ = InStr(II, aaa, SSS3)
    If JJ = 0 Then GoTo line_13040
    xxx = Left(aaa, JJ - 1) + s3_imbed + Mid(aaa, JJ + Len(SSS3) - 1)
    II = JJ + 1
    GoTo line_13032
'23 june 2002 add 4 thru 6 below
line_13040:
    If s4_imbed = "" Then GoTo line_13050
    II = 1
line_13042:
    JJ = InStr(II, aaa, SSS4)
    If JJ = 0 Then GoTo line_13050
    xxx = Left(aaa, JJ - 1) + s4_imbed + Mid(aaa, JJ + Len(SSS4) - 1)
    II = JJ + 1
    GoTo line_13042
line_13050:
    If s5_imbed = "" Then GoTo line_13060
    II = 1
line_13052:
    JJ = InStr(II, aaa, SSS5)
    If JJ = 0 Then GoTo line_13060
    xxx = Left(aaa, JJ - 1) + s5_imbed + Mid(aaa, JJ + Len(SSS5) - 1)
    II = JJ + 1
    GoTo line_13052
line_13060:
    If s6_imbed = "" Then GoTo line_13070
    II = 1
line_13062:
    JJ = InStr(II, aaa, SSS6)
    If JJ = 0 Then GoTo line_13070
    xxx = Left(aaa, JJ - 1) + s6_imbed + Mid(aaa, JJ + Len(SSS6) - 1)
    II = JJ + 1
    GoTo line_13062
line_13070:

line_13620:
    II = Len(aaa)
    array_pos = array_pos + 1
    If array_pos > 55 Then
'22 October 2004    xtemp = InputBox("error 13620 line too long=" + CStr(array_pos), "test", , xx1 - offset1, yy1 - offset2) '
        frmproj2.Caption = "error 13620 line too long (use crop option) " + CStr(Len(aaa)) '22 October 2004
            new_delay_sec = 1    '22 October 2004
            GoSub line_30300            '22 October 2004
        Beep                     '22 October 2004
        II = 50 '22 October 2004
        aaa = Left(aaa, 50) '22 October 2004
        array_pos = 10  '22 October 2004   keep it from stopping on line too long errors
      GoTo line_13635
    End If
line_13625:
    If II <= line_len + over_lap Then GoTo line_13630
line_13626:
    tt = InStr(line_len, xxx, " ")
    'if a - is going to split a word make sure that there are no nearby spaces below
line_13627:
    If (tt = 0 Or tt > line_len + over_lap) And Mid(aaa, line_len - 1, 1) = " " Then tt = line_len - 1
    If (tt = 0 Or tt > line_len + over_lap) And Mid(aaa, line_len - 2, 1) = " " Then tt = line_len - 2
    If (tt = 0 Or tt > line_len + over_lap) And Mid(aaa, line_len - 3, 1) = " " Then tt = line_len - 3
    If (tt = 0 Or tt > line_len + over_lap) And Mid(aaa, line_len - 4, 1) = " " Then tt = line_len - 4
line_13627a:
    If tt = 0 Then GoTo line_13628
line_13627b:
    If tt > line_len + over_lap Then GoTo line_13628
line_13627c:
    array_aaa(array_pos) = " " + Left(aaa, tt - 1) + " "
    array_ooo(array_pos) = " " + Left(ooo, tt - 1) + " "
    aaa = Mid(aaa, tt)
    ooo = Mid(ooo, tt)
    xxx = Mid(xxx, tt)  'january 21 2001
    GoTo line_13620
line_13628:
    new_len = line_len
    If match_flag = "NO" Then GoTo line_13628m
    'ensure the break does not come in any of the hi-lite items...
    'for each of the 3 search elements make sure there is no break
    'in the middle use s1len s2len s3len values and the instr(?)
    'to check for overlap
    
    'if the hi-lite item less than 2 characters the break can never split it.
'    new_len = line_len
    If s1len < 2 Then GoTo line_13628d
    III = line_len - s1len + 2
    If III < 1 Then GoTo line_13628d        'should never happen
    JJ = InStr(III, aaa, SSS1)
    If JJ = 0 Then GoTo line_13628d         'no match found
    If JJ > line_len Then GoTo line_13628d  'match found after break
'    If JJ + s1len - 1 = line_len Then GoTo line_13628d '
'   the above line was used when III above was calculated using II = line_len-s1len+1
    new_len = JJ - 1
    GoTo line_13628m
line_13628d:
    If s2len < 2 Then GoTo line_13628h
    III = line_len - s2len + 2
    If III < 1 Then GoTo line_13628h        'should never happen
    JJ = InStr(III, aaa, SSS2)
    If JJ = 0 Then GoTo line_13628h         'no match found
    If JJ > line_len Then GoTo line_13628h  'match found after break
    new_len = JJ - 1
    GoTo line_13628m
line_13628h:
    If s3len < 2 Then GoTo line_13628i
    III = line_len - s3len + 2
    If III < 1 Then GoTo line_13628i        'should never happen
    JJ = InStr(III, aaa, SSS3)
    If JJ = 0 Then GoTo line_13628i         'no match found
    If JJ > line_len Then GoTo line_13628i  'match found after break
    new_len = JJ - 1
    GoTo line_13628m
'23 june 2002 do 4 thru 6 below
line_13628i:
    If s4len < 2 Then GoTo line_13628j
    III = line_len - s4len + 2
    If III < 1 Then GoTo line_13628j        'should never happen
    JJ = InStr(III, aaa, SSS4)
    If JJ = 0 Then GoTo line_13628j         'no match found
    If JJ > line_len Then GoTo line_13628j  'match found after break
    new_len = JJ - 1
    GoTo line_13628m
line_13628j:
    If s5len < 2 Then GoTo line_13628k
    III = line_len - s5len + 2
    If III < 1 Then GoTo line_13628k        'should never happen
    JJ = InStr(III, aaa, SSS5)
    If JJ = 0 Then GoTo line_13628k         'no match found
    If JJ > line_len Then GoTo line_13628k  'match found after break
    new_len = JJ - 1
    GoTo line_13628m
line_13628k:
    If s6len < 2 Then GoTo line_13628m
    III = line_len - s6len + 2
    If III < 1 Then GoTo line_13628m        'should never happen
    JJ = InStr(III, aaa, SSS6)
    If JJ = 0 Then GoTo line_13628m         'no match found
    If JJ > line_len Then GoTo line_13628m  'match found after break
    new_len = JJ - 1

line_13628m:
    array_aaa(array_pos) = " " + Left(aaa, new_len) + "-"
    array_ooo(array_pos) = " " + Left(ooo, new_len) + "-"
    aaa = Mid(aaa, new_len + 1)
    ooo = Mid(ooo, new_len + 1)
    xxx = Mid(xxx, new_len + 1)
'
 '      xtemp = InputBox("testing prompt_=" + aaa + " " + CStr(new_len) + " " + CStr(cnt), "test", , 4400, 4500) '
'    xtemp = InputBox("testing prompt=" + AllSearch(1) + "*" + ttt + " " + SSS1 + " " + SSS2 + " " + SSS3 + " " + CStr(zzz_cnt), "test", , 4400, 4500) '
'     If UCase(xtemp) = "X" Then GoTo End_32000
    GoTo line_13620
'line_13629:
'    Print #ExtFile, Left(aaa, line_len); "-"
'    temp2 = temp2 + 1
'    aaa = Mid(aaa, line_len + 1)
'    GoTo line_13620
line_13630:
    array_aaa(array_pos) = " " + aaa + " "
    array_ooo(array_pos) = " " + ooo + " "
line_13635:
    If Left(array_aaa(1), 2) = "  " Then
        array_aaa(1) = Mid(array_aaa(1), 2) 'remove the extra space line 1 only
        array_ooo(1) = Mid(array_ooo(1), 2) 'as 1 is already added at start
    aaa = array_aaa(1)  'line broken into multiple parts start printing with first part
    ooo = array_ooo(1)
    array_prt = 1
    End If
line_13640:
    'back to the logic after 12120 line
    GoTo line_12120
line_13999:
line_14222:                 'january 15 2001
Return

line_14500:         'january 19 2001 check for imbedded spaces in search strings
    s1_imbed = ""
    s2_imbed = ""
    s3_imbed = ""
    s4_imbed = ""       '23 june 2002
    s5_imbed = ""
    s6_imbed = ""
    imbedded = "NO"
    If Len(SSS1) < 2 Then GoTo line_14510
    III = 2
line_14505:
    II = InStr(III, SSS1, " ")
    If II = 0 Or II = Len(SSS1) Then GoTo line_14510
    s1_imbed = Left(SSS1, II - 1) + "x" + Mid(SSS1, II + 1)
    III = II + 1
    GoTo line_14505
line_14510:
    If Len(SSS2) < 2 Then GoTo line_14520
    III = 2
line_14515:
    II = InStr(III, SSS2, " ")
    If II = 0 Or II = Len(SSS2) Then GoTo line_14520
    s2_imbed = Left(SSS2, II - 1) + "x" + Mid(SSS2, II + 1)
    III = II + 1
    GoTo line_14515
line_14520:
    If Len(SSS3) < 2 Then GoTo line_14530
    III = 2
line_14525:
    II = InStr(III, SSS3, " ")
    If II = 0 Or II = Len(SSS3) Then GoTo line_14530
    s3_imbed = Left(SSS3, II - 1) + "x" + Mid(SSS3, II + 1)
    III = II + 1
    GoTo line_14525
line_14530:
    If Len(SSS4) < 2 Then GoTo line_14540
    III = 2
line_14535:
    II = InStr(III, SSS4, " ")
    If II = 0 Or II = Len(SSS4) Then GoTo line_14540
    s4_imbed = Left(SSS4, II - 1) + "x" + Mid(SSS4, II + 1)
    III = II + 1
    GoTo line_14535
line_14540:
    If Len(SSS5) < 2 Then GoTo line_14550
    III = 2
line_14545:
    II = InStr(III, SSS5, " ")
    If II = 0 Or II = Len(SSS5) Then GoTo line_14550
    s5_imbed = Left(SSS5, II - 1) + "x" + Mid(SSS5, II + 1)
    III = II + 1
    GoTo line_14545
line_14550:
    If Len(SSS6) < 2 Then GoTo line_14560
    III = 2
line_14555:
    II = InStr(III, SSS6, " ")
    If II = 0 Or II = Len(SSS6) Then GoTo line_14560
    s6_imbed = Left(SSS6, II - 1) + "x" + Mid(SSS6, II + 1)
    III = II + 1
    GoTo line_14555
line_14560:
    If s1_imbed + s2_imbed + s3_imbed + s4_imbed + s5_imbed + s6_imbed > "" Then imbedded = "YES"
Return

Last_lines_15000:
    save_line = "15000"     'aug 08/99
'    frmproj2.Caption = program_info + " (cls #7)" '18Dec2013
    Cls                     '18 august 2003
'---------------------------------------------------- part a
'18 March 2003 ver=1.01 v
'    If Clear_Context_lines = Context_lines Then
        'clear the context records here **note the spaces=50 other setting = ""
    If Clear_Context_lines > 0 Then
'            tt1 = InputBox("testing=" + CStr(Context_cnt) + " " + CStr(Context_lines) + " " + CStr(Context_lines - Clear_Context_lines) + "=", , , 4400, 4500) 'TESTING ONLY
'        For II = 1 To MAX_CNT
    If Context_cnt > Context_lines - 1 Then
'        For II = 1 To Context_cnt - 1
        For II = 1 To Context_cnt - (Context_lines - Clear_Context_lines)
            Context_text(II) = Space(20)
        Next II
    Else
        For II = 1 To Context_cnt - (Context_lines - Clear_Context_lines)
            Context_text(II) = Space(20)
        Next II             'active testing this logic
        For II = Context_cnt + 1 To MAX_CNT Step 1
            Context_text(II) = Space(20)
        Next II
    End If
    End If
'    End If
'18 March 2003 ver=1.01 ^
'the above routine is handy when used in conjunction with a control file
'command type situation where the font size and clear_context_string cmd(55)
'element is used for display of text on top of existing photos
'may want a blank line or two right after as program will temporarily be
'switched to context display mode and may need to be swapped back???
'
'use the ver=1.01 changes when doing the display of the searched text
'on top of the picture on the screen in ver=1.02 "search criteria"
' "photo_detail" display on screen large font with own delay etc
'
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    
    If Context_cnt < Context_lines Then
        For temp4 = MAX_CNT - Context_lines + Context_cnt To MAX_CNT
        If Context_text(temp4) = "" Then GoTo line_15115    'january 22 2001
'april 10 2001       If InStr(1, Context_text(temp4), "append end") <> 0 Then
'       If InStr(1, Context_text(temp4), append_end1) <> 0 Then
      If InStr(1, UCase(Context_text(temp4)), UCase(append_end1)) <> 0 Then
         append_end1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_end1)), Len(append_end1))
        
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append end"
        SSS1 = append_end1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""   '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6
        
        Set_Fore = temp_fore
        GoTo line_15110
       End If              'hilite any append lines i put in
'april 10 2001       If InStr(1, Context_text(temp4), "append start") <> 0 Then
    ' with the ucase stuff below
     
      If InStr(1, UCase(Context_text(temp4)), UCase(append_start1)) <> 0 Then
         append_start1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_start1)), Len(append_start1))
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append start"
        SSS1 = append_start1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""   '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15110
       End If              'hilite any append lines i put in
       
       If InStr(1, Context_text(temp4), hilite_this) <> 0 And _
       (hilite_this <> "" And hilite_this <> "     ") Then
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = hilite_this
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001 this one
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4     '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15110
       End If              'hilite any append lines i put in

line_15100:
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = ""
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4     '23 june 2002
        SSS5 = new5
        SSS6 = new6

line_15110:
        cnt = cnt + 1
line_15115:             'january 22 2001
        Next temp4
'-------------------------------------------------------
        For temp4 = 1 To Context_cnt
        If Context_text(temp4) = "" Then GoTo line_15525    'january 22 2001

line_15200:
'april 10 2001       If InStr(1, Context_text(temp4), "append end") <> 0 Then
'       If InStr(1, Context_text(temp4), append_end1) <> 0 Then
      If InStr(1, UCase(Context_text(temp4)), UCase(append_end1)) <> 0 Then
         append_end1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_end1)), Len(append_end1))
        
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append end"
        SSS1 = append_end1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4     '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15520
       End If              'hilite any append lines i put in
'april 10 2001       If InStr(1, Context_text(temp4), "append start") <> 0 Then
'       If InStr(1, Context_text(temp4), append_start1) <> 0 Then
      If InStr(1, UCase(Context_text(temp4)), UCase(append_start1)) <> 0 Then
         append_start1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_start1)), Len(append_start1))
        
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4     '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append start"
        SSS1 = append_start1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15520
       End If              'hilite any append lines i put in
       
       If InStr(1, Context_text(temp4), hilite_this) <> 0 And _
       (hilite_this <> "" And hilite_this <> "     ") Then
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4         '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = hilite_this
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001 this one
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4     '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15520
       End If              'hilite any append lines i put in
        
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4         '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = ""
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""           '23 june 2002
        SSS5 = ""
        SSS6 = ""
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6

line_15520:
        cnt = cnt + 1
line_15525:             'january 22 2001
        Next temp4
    End If
    
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

    If Context_cnt > Context_lines - 1 Then
        For temp4 = Context_cnt - Context_lines + 1 To Context_cnt
        If Context_text(temp4) = "" Then GoTo line_15553    'january 22 2001
' get the context_text stuff to wrap vip
'april 10 2001       If InStr(1, Context_text(temp4), "append end") <> 0 Then
'      If InStr(1, Context_text(temp4), append_end1) <> 0 Then
      If InStr(1, UCase(Context_text(temp4)), UCase(append_end1)) <> 0 Then
         append_end1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_end1)), Len(append_end1))
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4         '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append end"
        SSS1 = append_end1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15550
       End If              'hilite any append lines i put in
'april 10 2001       If InStr(1, Context_text(temp4), "append start") <> 0 Then
'       If InStr(1, Context_text(temp4), append_start1) <> 0 Then
      If InStr(1, UCase(Context_text(temp4)), UCase(append_start1)) <> 0 Then
         append_start1 = Mid(Context_text(temp4), InStr(1, UCase(Context_text(temp4)), UCase(append_start1)), Len(append_start1))
        
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4             '23 june 2002
        new5 = SSS5
        new6 = SSS6
'april 10 2001        SSS1 = "append start"
        SSS1 = append_start1
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""       '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15550
       End If              'hilite any append lines i put in
       
       If InStr(1, Context_text(temp4), hilite_this) <> 0 And _
       (hilite_this <> "" And hilite_this <> "     ") Then
        tot_s1 = tot_s1 - 1         'january 06 2001
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4         '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = hilite_this
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""           '23 june 2002
        SSS5 = ""
        SSS6 = ""
        temp_fore = Set_Fore
        Set_Fore = AltColor
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        If hilite_hh = "Y" Then
            GoSub hilite_25500
        End If              'april 22 2001 this one
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6
        Set_Fore = temp_fore
        GoTo line_15550
       End If              'hilite any append lines i put in

line_15500:
        new1 = SSS1
        new2 = SSS2
        new3 = SSS3
        new4 = SSS4         '23 june 2002
        new5 = SSS5
        new6 = SSS6
        SSS1 = ""
        SSS2 = ""
        SSS3 = ""
        SSS4 = ""           '23 june 2002
        SSS5 = ""
        SSS6 = ""
        keep_aaa = aaa + ""
        keep_ooo = ooo + ""
        aaa = Context_text(temp4)
        ooo = Context_text(temp4)
        GoSub sub_12000
        aaa = keep_aaa
        ooo = keep_ooo
        SSS1 = new1
        SSS2 = new2
        SSS3 = new3
        SSS4 = new4         '23 june 2002
        SSS5 = new5
        SSS6 = new6
        
line_15550:
        cnt = cnt + 1
line_15553:             'january 22 2001
        Next temp4
    End If              'aug 08/99
line_15555:             'november 10 2000
'    Print Context_cnt   'november 10 2000 testing
    Context_cnt = -1     'november 10 2000
Return

line_16000:         'november 17 2000 enter file name
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 3 " + ttt, , , 4400, 4500) 'TESTING ONLY
'        tt1 = "replace.txt"         '27 October 2004 testing only reactivate line below
        Line Input #BatchFile, tt1
'            xtemp = InputBox("27 october 2004 TESTING 3a =" + tt1, , , 4400, 4500) 'TESTING ONLY
        GoTo skip_file
    End If                                              '27 October 2004
    Test1_str = "File name Prompt"              'april 01 2001
    If filereason <> "" Then Test1_str = filereason 'april 01 2001
'18Jun2012
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        tt1 = "MPEG_VIDEOS.TXT"
        GoTo skip_file
End If                  '18Jun2012
    tt1 = InputBox("***enter file name", Test1_str, xtemp, xx1 - offset1, yy1 - offset2) 'TESTING ONLY
skip_file:          '27 October 2004
    tt1 = UCase(tt1)
    If tt1 = "" Then
        FileExt = UCase(xtemp)
    Else
        FileExt = tt1
    End If
    If Right(FileExt, 4) <> ".TXT" Then
        GoTo line_16000
    End If
'            xtemp = InputBox("27 october 2004 TESTING 3b =" + tt1, , , 4400, 4500) 'TESTING ONLY
Return
line_16100:         'january 01 2001
'            save_line = "16100"
            save_line = "16100"
            Close ExtFile
'            Kill FileExt        'january 01 2001  january 22 2001
            DoEvents
            ExtFile = FreeFile
            ExtFile = ExtFile + 1   'maybe this will fix the problem?????
            DoEvents
'          save_line = "16100a"                '18 March 2007
    If oopen <> "APPEND" Then
     Open FileExt For Output Access Write As #ExtFile
    Else
     Open FileExt For Append Access Write As #ExtFile   'september 02 2001
    End If
    
'            save_line = "16100b"    '18 March 2007
            DoEvents
            line_len = 5000 'on extract do not do any wraps???
    'see the fix above re the adding of 1 to the file count with this prompt
Return

Display_pict_17000:                 'start of display_pict_17000 subroutine look for end
'        tt1 = InputBox("doug testing pics1", , , 4400, 4500)  'testing

    save_line = "17000"     'march 31/00
    If pp_entered = "YES" Then
        GoTo line_17045
    End If      'november 6 2000
    Line Input #OutFile, Next_line
    zzz_cnt = zzz_cnt + 1
    Next_line_save = Next_line
    Previous_line_save = Previous_line
    This_line = ooo

Line_17020:
'        Print "PICTURE FOUND="; This_line
 '       tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
 '       leave above as example for testing display and prompt
'    If InStr(UCase(Previous_line + This_line + Next_line), "XXX.")
    If InStr(UCase(Previous_line + Next_line), "XXX.") = 0 Then
        GoTo Line_17090
    End If      'no more xxx. to display
    'get the next line for searching for XXX.
    'Next_line_save = Next_line
    'clear Next_line if XXX. found there
    'Previous_line_save = Previous_line
    'clear Previous_line if XXX. found there
    'do logic that calls up to two paint program displays
    'here and check at end for Next_line
'switched the order here so that only 1 line of text
'will work ie photo bla bla bal then xxx.
'before it was picking up the one before it
    If InStr(UCase(Next_line), "XXX.") <> 0 Then
        Line_Search = Next_line
   '     tt1 = InputBox("next_line" + SSS1 + Next_line, , , 4400, 4500) 'TESTING ONLY
    '    tt1 = InputBox("Previous_line" + SSS2 + Previous_line, , , 4400, 4500) 'TESTING ONLY
        Next_line = "" 'so we don't do it again
        Previous_line = ""      'line added
        GoTo Line_17025         'line added
    End If      'no more xxx. to display
    If InStr(UCase(Previous_line), "XXX.") <> 0 Then
        Line_Search = Previous_line
        Previous_line = "" 'so we don't do it again
        'goto Line_17025        'line removed
    End If      'no more xxx. to display

Line_17025:
    String_Position = InStr(UCase(Line_Search), "XXX.")
    If String_Position = 0 Then
        GoTo Line_17090:
    End If
        Pict_file = Right(Line_Search, Len(Line_Search) - String_Position - 3)
        Save_file = Pict_file   '15 January 2004
'30 May 2003 allow for cmd(58) DEFAULT_TO_CD to over-ride path... ver=1.06
    If UCase(Left(Cmd(58), 13)) = "DEFAULT_TO_CD" And Left(App.Path, 3) <> "C:\" Then
        Pict_file = Left(App.Path, 3) + Right(Pict_file, Len(Pict_file) - 3)
    End If          'change any C:\ to D:\ or E:\ etc 30 May 2003 ver=1.06
    
    If auto_exe = "YES" And Left(App.Path, 3) <> "C:\" Then
        Pict_file = Left(App.Path, 3) + Right(Pict_file, Len(Pict_file) - 3)
    End If          'change any C:\ to D:\ or E:\ etc 07 december 2002
line_17030:
    tt = InStr(Pict_file, Chr(9)) 'check for tabs
    If tt = 0 Then
        GoTo line_17035
    End If
    'change any tabs to 4 spaces
    Pict_file = Left(Pict_file, tt - 1) + "    " + Right(Pict_file, Len(Pict_file) - tt)
    GoTo line_17030
line_17035:
    If Left(Pict_file, 1) <> " " Then
        GoTo line_17040     'STRIP OUT LEADING SPACES
    End If
    Pict_file = Right(Pict_file, Len(Pict_file) - 1)
    GoTo line_17035
line_17040:
    'check the file exists before going anywhere
    save_line = "17040"
    FileFile = FreeFile
'13 December 2004 had to comment out the following 2 lines to get the error trap on mcisendstring open to work
'13 December 2004    Open Pict_file For Input As #FileFile
'13 December 2004    Close FileFile
'13 December 2004 maybe check to see what it would take to just check mpg or some such above ***vip*** todo
    DoEvents
    save_line = "17040a"        '13 December 2004
'        tt1 = InputBox("doug testing pics2 " + Line_Search, , , 4400, 4500) 'testing 29 november 2006
' aug 08/00
    
    If InStr(UCase(Line_Search), ".BMP") <> 0 And _
       Test1_str = "P1" Then
        GoTo line_17045
    End If
    avi_file = "NO"             '01 february 2003
    mpg_file = "NO"             '11 May 2003 ver=1.05
    wav_file = "NO"             '10 June 2003 ver=1.07
    mid_file = "NO"             '10 June 2003 ver=1.07
    If InStr(UCase(Line_Search), ".AVI") <> 0 Then
        avi_file = "YES"
        GoTo line_17045         '01 february 2003
    End If                      '01 february 2003
    
    If Mid(UCase(Line_Search), 5, Len(Line_Search) - 4) = "HTTP:" Then
'        tt1 = InputBox("doug testing pics2 " + Line_Search, , , 4400, 4500) 'testing 29 november 2006
        GoTo Line_17080         '
    End If                      '29 November 2006
   
    
'10 June 2003 ver=1.07    If InStr(UCase(Line_Search), ".MPG") <> 0 Then
'16 August 2004    If InStr(UCase(Line_Search), ".MPG") <> 0 Or InStr(UCase(Line_Search), ".MP3") Then
    If InStr(UCase(Line_Search), ".MPG") <> 0 Or InStr(UCase(Line_Search), ".MP3") Or InStr(UCase(Line_Search), ".WMV") Or InStr(UCase(Line_Search), ".VOB") Then
        mpg_file = "YES"
        GoTo line_17045         '11 May 2003
    End If
    
    'ver=1.07
    If InStr(UCase(Line_Search), ".WAV") <> 0 Then
        wav_file = "YES"
        GoTo line_17045         '10 June 2003
    End If
    
    'ver=1.07
    If InStr(UCase(Line_Search), ".MID") <> 0 Then
        mid_file = "YES"
        GoTo line_17045         '10 June 2003
    End If
    

'28 November 2004    If InStr(UCase(Line_Search), ".JPG") = 0 Then
    If InStr(UCase(Line_Search), ".JPG") = 0 And InStr(UCase(Line_Search), ".GIF") = 0 Then
        GoTo Line_17080
    End If
    
line_17045:
'28 November 2004 gif here too along with the jpg
'        tt1 = InputBox("doug testing pics2a " + Pict_file, , , 4400, 4500) 'testing
'            frmproj2.Caption = " testing=" + Pict_file + "=" 'testing
'only .JPG files below see line_17080 and gif now  28 november 2004
        LastFile = Pict_file
    'check the program exists before going anywhere
    save_line = "17041"
    FileFile = FreeFile
line_17050:
    save_line = "17050"
    Close FileFile
    DoEvents
'        tt1 = InputBox("doug testing pics2.1", , , 4400, 4500)  'TESTING ONLY
'january 03 2002 the nt computer at work failed on the following open
'why oh why are we doing the open cmd(8) below ***vip*** todo check this out
'they should all have explorer anyway...
    If debug_photo Then         '12 october 2002
            tt1 = InputBox("testing photo 6b*", , , 4400, 4500)  'TESTING ONLY
    End If
'09 september 2003    If Left(os_ver, 10) = "Windows 20" And Len(os_ver) > 15 Then
    If Left(os_ver, 10) = "Windows 20" Then
'09 september 2003        Open Cmd(43) For Input As #FileFile
'09 september 2003    Else
'10 September 2003        Open Cmd(8) For Input As #FileFile
    End If          '30 november 2002
    
'10 September 2003    Close FileFile
    save_line = "17055-1"     '16Aug2016 why is this not at line 17055  maybe test fixing this
        ' testing below explorer
'the P1 option should work for .bmp .jpg and .tif
'        tt1 = InputBox("doug testing pics2.2", , , 4400, 4500)  'TESTING ONLY
    'iffy 17050
    If (InStr(UCase(Cmd(8)), "EXPLORER") <> 0 Or InStr(UCase(Cmd(8)), "NETSCAPE") <> 0) And Test1_str = "P1" Then
'        Print "file name="; Pict_file; "="
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
'question? why the reference to explorer only to make
'question? the above p1 option work.
'        tt1 = InputBox("doug testing pics2.3", , , 4400, 4500)  'TESTING ONLY
'20 May 2003        Print Pict_file      'there was no need for this meebee
'    Heightnew = 10000   'february 16 2001 the new has no effect
'    Widthnew = 12000     'february 16 2001
'    Height = 9000       'february 16 2001 this changes the size of the form
'    Width = 10500        'february 16 2001
'demo the above using the p1 option and any picture

'february 21 2001        Set Picture = LoadPicture(Pict_file)

'  P I C T U R E S   D I S P L A Y E D   B E L O W

'most of the picture displays take place right below here with the load picture **vip**
'        tt1 = InputBox("doug testing pics3", , , 4400, 4500)  'TESTING ONLY
    If debug_photo Then         '12 october 2002
            tt1 = InputBox("testing photo 6b** " + tempdata + " " + motion_yn + " " + Pict_file, , , 4400, 4500) 'TESTING ONLY
    End If
    
'29 March 2003      ver=1.04
    motion_yn = "NO"        '29 March 2003
    
'28 february 2003       skip dups if there
'handy to see the duplicates if a visual inventory otherwise skip em for most part
'19 September 2004 do Command== stuff here
'21 September 2004 might want to check for command== right after the input line.****************
'might need to keep the zzz_cnt number and on return read 1 past that count ie reopen...
'19 November allow for switch of control files...
'***vip*** todo maybe allow it to change input files here too
'          to whatever is in the corresponding fields in cmd(46) cmd(47) cmd(48) etc
eof_entry:                                  '26 November 2004
'            tt1 = InputBox("testingy " + aaa + " " + Cmd(73) + " " + TheFile, , , 4400, 4500) 'TESTING ONLY 26 NOvember 2004
    III = InStr(UCase(aaa), "CONTROL==")    '19 November 2004
    If III <> 0 Then
'            tt1 = InputBox("testingA " + control_file + " " + aaa + " " + Cmd(73), , , 4400, 4500) 'TESTING ONLY 26 NOvember 2004
        control_file = Trim(Right(aaa, Len(aaa) - III - 8)) '19 November 2004
            tempdata = Cmd(73)                      '22 November 2004  the fileswitch info
'        control_file = "chgfile.txt"           '26 November 2004 testing with this
'            frmproj2.Caption = " testing=" + control_file + "=" '26 November 2004
        GoSub Control_28000     'change the control file from in a *.txt file
        DoEvents
        If Left(UCase(tempdata), 10) = "FILESWITCH" Then
            Close #OutFile          '22 November 2004
            DoEvents
            OutFile = FreeFile
            DoEvents
            TheFile = Trim(Cmd(46))  'Must be a file name here not a number
'            SSS1 = "A"          '26 November 2004 testing this
'            frmproj2.Caption = " testing=" + TheFile + "=" + SSS1 '26 November 2004
'            tt1 = InputBox("testingx " + control_file + " " + Cmd(46) + " " + TheFile, , , 4400, 4500) 'TESTING ONLY 26 NOvember 2004
            Open TheFile For Input As #OutFile
'            tt1 = InputBox("testingx " + control_file + " " + Cmd(46) + " " + TheFile, , , 4400, 4500) 'TESTING ONLY 26 NOvember 2004
            DoEvents
'===========================
'            frmproj2.Caption = " testingx=" + TheFile + "=" + Test2_str
            
'            tt1 = InputBox("testingX " + control_file + " " + Cmd(47) + " " + TheFile + " " + Cmd(48), , , 4400, 4500) 'TESTING ONLY 26 NOvember 2004
            tempdata = "CONTROL"         '26 November 2004       allow for skip of prompt three with this
            multi_prompt2 = RTrim(Cmd(47))       '26 NOVEMBER 2004
            '18 December 2004 the following line did a fix for the file switching...
            Cmd(73) = "noFILESWITCH"        '18 December 2004 if the new file had FILESWITCH set it would go wild
            temptemp = Cmd(48)              '26 November 2004  the search string ie "A" or "PHOTO"
            If eofsw = "YES" Then
                eofsw = ""
                ttt = multi_prompt2
            End If
            GoTo Get_no2                        '26 November 2004
'===========================
        End If                      '22 November 2004
'            frmproj2.Caption = " testing1x=" + control_file + "=" '26 November 2004
        GoTo input_1000a        '19 November 2004  so the pause does not take place etc etc.
    End If                                      '19 November 2004
    If Len(command_line) < 1 Then
        command_line = ""
        III = InStr(UCase(aaa), " COMMAND==")
        If III <> 0 Then
            II = InStr(III + 9, aaa + " ", " ") 'make sure there is a trailing space here
            command_line = Mid(aaa, III + 10, II - III - 10)
            multi_prompt2 = command_line
            hold_zzz = zzz_cnt          '21 September 2004
'            aaa = Left(aaa, III + 1) + Right(aaa, Len(aaa) - II - 1) '21 September 2004
            If SSS1 <> "" Then temptemp = SSS1  '02 January 2005
            If SSS2 <> "" Then temptemp = temptemp + sep + SSS2 '02 January 2005
            If SSS3 <> "" Then temptemp = temptemp + sep + SSS3 '02 January 2005
            If SSS4 <> "" Then temptemp = temptemp + sep + SSS4 '02 January 2005
            If SSS5 <> "" Then temptemp = temptemp + sep + SSS5 '02 January 2005
            If SSS6 <> "" Then temptemp = temptemp + sep + SSS6 '02 January 2005
'        tempdata = InputBox("02 January 2005 test=" + temptemp + " " + SSS1 + " " + SSS2 + " " + SSS3, CStr(zzz_cnt), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
            tempdata = "COMMAND"         '26 November 2004       allow for skip of prompt three with this
            
'02 January 2005            temptemp = Cmd(48)           '26 November 2004
'01 January 2005 the above line results in the search prompt being "photo" change this to keep the old prompt
'ie screen/saver  etc etc
            GoTo Get_no2
        End If                      '19 September 2004
    Else                    '21 September 2004
'            command_line = ""   '21 September 2004
    End If                      '21 September 2004
If Pict_file = last_pict And Cmd(54) <> "SHOWDUPS" Then
    GoTo Line_17055         '28 february 2003
End If                      '28 february 2003
'12 September 2004 check if "backgrd==" in the line here and fire up the background job...
'12 September 2004
        back_job = ""
    save_line = "17055-2"     '23Aug2016
        III = InStr(UCase(aaa), " BACKGRD==")
        If III <> 0 Then
            II = InStr(III + 9, aaa + " ", " ") 'make sure there is a trailing space here
            back_job = Mid(aaa, III + 10, II - III - 10)
'        xtemp = InputBox("12 Sept test=" + App.Path + "\" + back_job, CStr(delay_sec), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
            If InStr(1, back_job, ":") = 0 Then
                back_job = App.Path + "\" + back_job
            End If
'02 October 2004 check the back_job here
'    frmproj2.Caption = "before test " + back_job     '02 October 2004 testing
'            new_delay_sec = 2    '02 October 2004
'            GoSub line_30300            '02 October 2004
         MyAppID = Shell(back_job, 3)
         SendKeys "^o", True
        End If                      '12 September 2004

'05 march 2003              do a mp3 play
'    Pict_file = "E:\Dan\mp3_0001.mp3"  'testing only 20 march 2003 commented it out
'11 June 2003 ver=1.07 If InStr(1, UCase(Pict_file), ".MP3") <> 0 Then
'            frmproj2.Caption = "30Jun2012a Previous_line=" + Previous_line + " mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " sscreen_saver_ww=" + sscreen_saver_ww
'            xtemp = InputBox(" 30Jun2012a testing prompt", , , 4400, 4500)  'TESTING ONLY
'            frmproj2.Caption = "30Jun2012a Pict_file=" + Pict_file + " line_search=" + Line_Search
'            xtemp = InputBox(" 30Jun2012a testing prompt", , , 4400, 4500)  'TESTING ONLY
    save_line = "17055-3"     '23Aug2016
If InStr(1, UCase(Pict_file), ".MP3") <> 0 And motion_yn = "xxxx" Then '11 June 2003 deactivated this logic for now
    motion_yn = "YES"       '29 March 2003
Last$ = frmproj2.hWnd & " Style " & &H40000000
' ToDo$ = "open e:\cottonwood\101-0186_mvi.avi Type avivideo Alias video1 parent " & Last$
 todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 , vbnullstring, 0& ,0& "
            tt1 = InputBox("testing mp3 " + Pict_file, , , 4400, 4500) 'TESTING ONLY
 i = mciSendString("close video1", 0&, 0, 0)
 i = mciSendString(todo$, 0&, 0&, 0&)
' i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
' i = mciSendString("play video1 wait", 0&, 0, 0)
            tt1 = InputBox("testing mp3 a" + Pict_file, , , 4400, 4500) 'TESTING ONLY
 i = mciSendString("play video1", vbNullString, 0, 0)
 i = mciSendString("close video1", 0&, 0, 0)
     save_line = "17055-4"     '23Aug2016
           tt1 = InputBox("testing mp3 b" + Pict_file, , , 4400, 4500) 'TESTING ONLY
    GoTo Line_17055
End If                      '05 March 2003
    save_line = "17055-31"     '23Aug2016


'19 january 2003            avi file display here
If avi_file = "YES" Then
    motion_yn = "YES"       '29 March 2003
    If videoyn = "SHOWVIDEO" Then       '10 february 2003
Last$ = frmproj2.hWnd & " Style " & &H40000000
' ToDo$ = "open e:\cottonwood\101-0186_mvi.avi Type avivideo Alias video1 parent " & Last$
 todo$ = "open " + Pict_file + " Type avivideo Alias video1 parent " & Last$
 i = mciSendString(todo$, 0&, 0, 0)
' i = mciSendString("put video1 window at 16 10 124 120", 0&, 0, 0)
' i = mciSendString("put video1 window at 16 10 800 600", 0&, 0, 0)
' i = mciSendString("put video1 window at 16 10 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
 i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
' i = mciSendString("put video1 window at 16 10 1084 800", 0&, 0, 0)    'for 17 inch screen
 i = mciSendString("play video1 wait", 0&, 0, 0)
 i = mciSendString("close video1", 0&, 0, 0)
'19 january 2003 commented out till ready
    End If                  '10 february 2003
    GoTo Line_17055         '01 february 2003
End If                      '01 february 2003 end of avi_file = "YES"
    
'11 May 2003            mpg file display here ver=1.05
    save_line = "17055-32"     '23Aug2016
If mpg_file = "YES" Then
    motion_yn = "YES"
    slomo = False       '08 January 2004
    slomo = True        '31Dec2011

  If videoyn = "SHOWVIDEO" Then
    DoEvents        '17Jan2012
    Last$ = frmproj2.hWnd & " Style " & &H40000000
' ToDo$ = "open e:\cottonwood\101-0186_mvi.avi Type avivideo Alias video1 parent " & Last$
    '18 june 2003 re use short name as long ones with spaces cause a problem
    'might need to use this elsewhere ***vip*** todo
'17Jan2012 maybe comment out the stuff below trying the other name
'17Jan2012    lresult = getshortpathname(Pict_file, sshortfile, Len(sshortfile))
'17Jan2012    long_pict_file = Pict_file + ""            '22 November 2006
'17Jan2012    Pict_file = Left$(sshortfile, lresult)  '18 june 2003 needed the short file name
'11Feb2012 try taking out this stuff
    lresult = getshortpathname(Pict_file, sshortfile, Len(sshortfile))
    long_pict_file = Pict_file + ""            '22 November 2006
    Pict_file = Left$(sshortfile, lresult)  '18 june 2003 needed the short file name check this out 16Aug2016

'17Jan2012    If elapse_yn = "YES" Then elapse_start = Timer            '13 July 2003
'17Jan2012 todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 parent " & Last$
    save_line = "17055-33"     '23Aug2016
' todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 parent " & Last$
 todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 wait parent " & Last$ '03Mar2018
    save_line = "17055-33a"     '28Aug2016
' todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 wait parent " & Last$   '22Feb2012 test the wait
'11Feb2012        lresult = getshortpathname(Pict_file, sshortfile, Len(sshortfile))  '04Feb2012
'11Feb2012        Pict_file = Left$(sshortfile, lresult)                              '04Feb2012
'11Feb2012 todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 wait parent " & Last$
'   For II = 1 To 50
'        DoEvents
 '       DoEvents        '17Jan2012
 '   Next II         '17Jan2012 maybe this will help giving it some time
'18 March 2004 i = mciSendString("close all", 0&, 0, 0)      '17 March 2004 make sure nothing is there.

    '12 December 2004 this does not seem to trap any error but microsoft can check out the open command
'17Jan2012 maybe this is the error if I hit enter before the file starts it seems OK what the hey
'14Apr2012        frmproj2.Caption = "(Error Opening Video File)1 " + Save_file '12 December 2004 bad files stop here....
        frmproj2.Caption = "(Opening mpeg File1) " + Save_file '12 December 2004 bad files stop here....
'    Timer1.Enabled = True '13 December 2004
'    Timer1.Enabled = False '13 December 2004
'    Timer1_Timer          '13 December 2004 all the timer
' i = mciSendString(todo$, 0&, 0, 0)         'open command
'28Jan2012 need to skip the open here if continueyn = "Y"
'30Jan2012
' testtest = " delay_sec=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo)  '17Jan2012 teststring
'       testtest = InputBox("continueyn= " + continueyn + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
    If pauseyn <> "Y" And continueyn <> "Y" Then
                    i = mciSendString("close video1", 0&, 0, 0)
              '      pauseyn = "N"
              '      hold_pauseyn = ""
    End If          '25Mar2012
'20Feb2018a  test prompt here...
'    tt1 = InputBox("test03Mar2018b=" + xtemp + "=" + " before video len call " + " " + CStr(errormsg1), , , 4400, 4500) '20Feb2018
    If continueyn = "Y" Then GoTo no_open       '28Jan2012
    save_line = "17055-34"     '23Aug2016
 result1 = mciSendString(todo$, returnstring1, 1024, 0)   '12 December 2004a
    save_line = "17055-35"     '23Aug2016
If Not result1 = 0 Then
    errormsg1 = mciGetErrorString(result1, errorstring1, 1024) '12 December 2004a
        frmproj2.Caption = " err  =" + CStr(result1) + " at " + save_line + " opening video file 3" + Save_file + " " + errorstring1 '12 December 2004
End If  '03Mar2018 display any error
'    tt1 = InputBox("test03Mar2018b=" + xtemp + "=" + " after open video call " + " " + CStr(result1), , , 4400, 4500) '20Feb2018
'10Feb2012 testing only below This group of 11Feb2012 displays were used to debug the fast forward option
'20Feb2018 check this out maybe
'***vip*** just comment them out and use at another time it seems the laptop needs more play time to work than the pc
' testing = "video_audit1=" + mpg_file + " returnstring1=" + returnstring1 + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'        frmproj2.Caption = "check the alias =" + video1 '12 December 2004
'            new_delay_sec = 2
'            GoSub line_30300        '13 December 2004
'    Timer1.Enabled = False '13 December 2004
'    Timer1.Enabled = True '13 December 2004
'    On Error GoTo 0             '12 December 2004a
    DoEvents
'13 December 2004 try and put the following command in a timer
'13 December 2004 result = mciSendString(todo$, ByVal 0&, 0, 0)   '12 December 2004a
'        xtemp = InputBox("13 December 2004 test erro test=", CStr(delay_sec), , xx1 - offset1, yy1 - offset2)   '29 February 2004 test
'14Apr2012        frmproj2.Caption = "error opening video file2 ="  '12 December 2004
        frmproj2.Caption = "opening video file2 ="  '12 December 2004
    DoEvents
'    tt1 = InputBox("test03Mar2018c=" + xtemp + "=" + CStr(video_length) + " " + CStr(errormsg1), , , 4400, 4500) '20Feb2018
If Not result1 = 0 Then
    errormsg1 = mciGetErrorString(result1, errorstring1, 1024) '12 December 2004a
        frmproj2.Caption = " err  =" + CStr(result1) + " at " + save_line + " opening video file 3" + Save_file + " " + errorstring1 '12 December 2004
'            new_delay_sec = 2       '12Aug2016 no need for this delay
'12Aug2016            GoSub line_30300        '13 December 2004
'            new_delay_sec = 2       '12Aug2016 no need for this delay
'            GoSub line_30300        '13 December 2004
    save_line = "17055-35"     '23Aug2016
        If (result1 = 277 Or result1 = 263) And copy_photo = "YES" And Left(save_line, 5) = "17055" Then
            'need to switch the output file right here badmp3 is captured here
            photo_dir = badmp3_fold      '12Aug2016
'            GoTo line_2153 '25Aug2016
'             new_delay_sec = 2       '12Aug2016 no need for this delay
'            GoSub line_30300        '13 December 2004
            GoTo no_open        '25Aug2016
        End If          '12Aug2016
        GoTo input_1000             '12 December 2004
 End If                         '12 December 2004
'            On Error GoTo Errors_31000      '12 December 2004a
    save_line = "17055-36"     '23Aug2016
 i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
    save_line = "17055-37"     '23Aug2016
'23Mar2012 testing only below
' testing = "video_audit2=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
no_open:            '28Jan2012
'28Jan2012 maybe skip to here if continueyn = "Y"
'26 April 2004 try and start getting output data
' i = mciSendString("set video2 channels 1 ", 0&, 0, 0)  'new
 'todo$ = "open new c:\search\outmpg.mpg" + " Type MPEGVideo Alias video2 parent " & Last$
 'i = mciSendString(todo$, 0&, 0, 0)
' i = mciSendString("video2 video record", 0&, 0, 0) '26 February 2004
'26 April 2004
 
 '26 February 2004  find the length of the video file here
'01 April 2004 i = mciSendString("status video1 length", mssg, 255, 0) '26 February 2004
'19 August 2004 try and mute the video if not full speed
'the following mute made no difference.
'the slow speed seems to mute it somewhat by itself.
If slomo = True Then
'17Jan2012 could be something here that needs doing
'31Dec2011 deactivate the mute for now i = mciSendString("set video1 audio all off", 0&, 0, 0)
'             SendKeys "^c"                   '01 March 2007 check this out more
' i = mciSendString("set video1 audio right off", 0&, 0, 0)
' i = mciSendString("set video1 audio left off", 0&, 0, 0)
End If          '19 August 2004 try to mute video if slomo
'***vip*** need to see why this logic below does not work too well skipping the open is important for fast forward etc
'20Feb2012    If continueyn = "Y" Or hold_pauseyn = "Y" Then GoTo no_lenreqd  '20Feb2012
    If continueyn = "Y" Then GoTo no_lenreqd  '20Feb2012
'20feb2012    If (continueyn = "Y" And save_play_speed = 1000) Or (pauseyn = "Y" And save_play_speed = 1000) Then GoTo no_lenreqd '18Feb2012
'    If Not rand1 Then GoTo no_lenreqd        '11Feb2012
'12 December 2004 maybe the video stops here...
    save_line = "17055-wait"     '20Feb2018a
        frmproj2.Caption = "(Finding Video length)"  '12 December 2004
get_lenagain:       '20Feb2018
 i = mciSendString("status video1 length wait", mssg, 255, 0) '26 February 2004
' i = mciSendString("status video1 length ", mssg, 255, 0) '20Feb2018a

' i = mciSendString("status video1 length wait", mssg, 255, 0) '20Feb2018 tried multiples of this line no workie nope...
    hold_len = mssg + ""    '20Feb2012
 temp3 = InStr(mssg, Chr$(0)) '26 February 2004
 video_length = Val(Left(mssg, temp3 - 1)) '26 February 2004
If Not i = 0 Then
    errormsg1 = mciGetErrorString(i, errorstring1, 1024) '20Feb2018
        frmproj2.Caption = " err  =" + CStr(result1) + " at " + save_line + " opening video file 3" + Save_file + " " + errorstring1 '12 December 2004
End If                         '20Feb2017
'    If debug_photo = True Then
'        tt1 = InputBox("test20Feb2018=" + xtemp + "=" + CStr(video_length) + " " + CStr(errormsg1), , , 4400, 4500) '20Feb2018a
'    End If  '20Feb2018a
'    tt1 = InputBox("test20Feb2018x=" + xtemp + "=" + CStr(video_length) + " " + CStr(errormsg1), , , 4400, 4500) '20Feb2018
    If tt1 = "x" Then GoTo get_lenagain     '20Feb2018a
no_lenreqd:        '20feb2012
    mssg = hold_len + "" '20Feb2012
 temp3 = InStr(mssg, Chr$(0)) '26 February 2004
 video_length = Val(Left(mssg, temp3 - 1)) '26 February 2004
' video_length = 262260 '20Feb2018 see if forcing it works (nope)
'        tt1 = InputBox("test20Feb2018=" + xtemp + "=" + CStr(video_length), , , 4400, 4500) '20Feb2018
'22Nov2010
    If delay_sec < 0# Then
        delay_sec = video_length / 1000 + delay_sec
        line_delay_sec = delay_sec
'        tt1 = InputBox("testing#2=" + xtemp + "=" + CStr(delay_sec), , , 4400, 4500) '21Jan2012
    End If                          '22Nov2010
'25Jun2010 see video_length above need that less the 3 seconds to determine start== value
'23 March 2004      do random start point here
    '28 March 2004 some files have bad size data ie demo76_pict6.mpg
'20Feb2018 video_length = 262260
    If video_length > 50000000 Then '28 March 2004 ie the value is 19431083 ??
        frmproj2.Caption = "(Bad Video length=" + CStr(video_length) + ")" '28 March 2004
            new_delay_sec = 4
            GoSub line_30300        '28 March 2004
        video_length = 10000        'on a bad number set to 10 seconds long
    End If                          '28 March 2004
    If rand1 Then
        rand_cnt1 = video_length - (hold_sec * 1000)    '23 March 2004
        rand_no1 = Int(rand_cnt1 * Rnd + 1)   '23 March 2004
        line_start_point = rand_no1             '23 March 2004
check_l_s_p:                    '27 March 2004
        If line_start_point > 300000 Then
            line_start_point = line_start_point - 300000
            GoTo check_l_s_p
        End If                  '27 March 2004
    keep_line_start_point = line_start_point    '25Jan2012
    keep_begin_point = line_start_point         '25Jan2012
    begin_point = line_start_point         '25Jan2012
    
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
' Test2_str = InputBox("testing for thumb_nail= " + thumb_nail + " " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , 4000, 5000) '
                    '27 March 2004 jumping to the middle of large files takes too line
                    'this limits that to the first 5 minutes and that is plenty.
                    
    End If                                      '23 March 2004

'29 February 2004 leap year extra day
'14 March 2004 do not allow the delay to be longer than the file below ie 83/3/7 example
'        xtemp = InputBox("24 March 2004 test length=" + CStr(video_length) + " " + CStr(line_delay_sec) + " " + CStr(line_start_point), CStr(delay_sec), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
'    If thumb_nail = "YES" Then line_delay_sec = hold_sec       '22 March 2004
    If thumb_nail = "YES" Then
        line_delay_sec = hold_sec       '25Jan2012
        keep_line_delay_sec = hold_sec       '25Jan2012
    End If                              '25Jan2012
'25Jan2012
' testtest = resume_str + "mpg_file=" + mpg_file + " picture_search=" + Picture_Search + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " picture_search=" + Picture_Search + " line_start_point=" + CStr(line_start_point) + " begin_point=" + CStr(begin_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " delay_sec=" + CStr(delay_sec) + " " '08Nov2011 teststring
' Test2_str = InputBox("testing for thumb_nail= " + thumb_nail + " " + testtest, "Interrupt Prompt # 0  " + CStr(dsp_cnt) + " " + save_line, , 4000, 5000) '
'25 October 2004 remove the link from line_delay_sec below add new command element
'so the short small tiny video file length will now display (ie tt.01) FIXES
'how did line_start_point below ever go negative ????
    If line_start_point < -0.5 Then line_start_point = video_length + line_start_point  '25Jun2010
    If line_start_point < 0# Then line_start_point = 10        '25 October 2004 this did the fix on the short ones.
'25Jun2010 might want to check the line below for the very short vids???
'25Jun2010 comment this out    If line_start_point < 9.0001 Then line_start_point = 10    '11 March 2007
'        xtemp = InputBox("25 oct test video_length=" + CStr(video_length), CStr(line_delay_sec), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
'11 March 2007    If start_point <> 0 And rand1 <> -1 Then
    If start_point <> 10 And rand1 <> -1 Then
        line_start_point = start_point                  '20 August 2005
    End If
    If line_delay_sec * 1000 + line_start_point > video_length - 100 Then
'20 August 2005        line_delay_sec = 0              '14 March 2004
        line_start_point = video_length - line_delay_sec * 1000 '20 August 2005
'        xtemp = InputBox("20 August 2005 line_delay_sec=" + CStr(line_delay_sec) + " " + CStr(line_start_point), CStr(delay_sec), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
    End If
    If line_delay_sec < 0.0001 Then
        line_delay_sec = (video_length - line_start_point) / 1000
'        xtemp = InputBox("25 oct line_delay_sec=" + CStr(line_delay_sec) + " " + CStr(line_start_point), CStr(delay_sec), , xx1 - offset1, yy1 - offset2) '29 February 2004 test
        delay_sec = line_delay_sec
    End If          '29 February 2004 now that I have the length use it.
    If line_delay_sec > delay_sec Then
        delay_sec = line_delay_sec
    End If
    last_vs = 0             '25 March 2004
'        xtemp = InputBox("length=", CStr(video_length), , xx1 - offset1, yy1 - offset2)  '26 February 2004 test
'    If delay_sec = 11111 Then delay_sec = video_length / 1000       '21Jan2012
    '25Jan2012
    If thumb_nail = "YES" Then
        delay_sec = hold_sec       '25Jan2012
    End If                              '25Jan2012
    If delay_sec = 11111 Then
        delay_sec = video_length / 1000       '23Jan2012
        line_delay_sec = delay_sec          '23Jan2012
        temp_double = delay_sec
    End If                                  '23Jan2012
' testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secY=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info XX " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
' testtest = CStr(video_length) + " delay_sec1=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
 
'    start_point = 5000      'testing 16 july 2003 see if it works here first
'16 July 2003 (why just mp3 do both) -deactivated- If InStr(UCase(Line_Search), ".MP3") <> 0 Then
 If InStr(UCase(Line_Search), ".MPx") <> 0 Then
    i = mciSendString("set mp3 time format trnsf", 0&, 0, 0)
    '18 june 2003 testing
'    i = mciSendString("play " & Pict_file, 0&, 0&, 0&)
'16 July 2003    i = mciSendString("play video1 to 10000000", 0&, 0&, 0&)
    i = mciSendString("play video1 from " + CStr(start_point), 0&, 0&, 0&)
        'see about changing the above with start= and end= data
'    mediaplayer1.Open Pict_file
'    i = mciSendString("play video1 from 0 to 111111", 0&, 0&, 0&)
 End If             '11 june 2003 ver=1.07
'20Feb2012 no_lenreqd:             '11Feb2012

'********************
'14 July 2003 this is where a wait nowait check needs to be done below (see notify / wait options)
' i = mciSendString("play video1 wait", 0&, 0, 0)    'this is where the display happens
    temp_double = start_point           '19 July 2003 Ver=1.07T
'    play_speed = 150         '29 October 2003        (move this later)
'30 October de-activate the speed command below It seems to work on the laptop
'   but freezes the other two big screens set it to 500 for half and 1000 for full speed
'   The search.txt file gets corrupted and needs replacing  30 October 2003
'   the control.txt file gets corrupted needs some additional spaces? 30 October 2003
    slomo = False       '08 January 2004
    slomo = True        '31Dec2011
'    frmproj2.Caption = program_info + " Video File= " + Save_file '15 January 2004
'    frmproj2.Caption = program_info + " Video info= " + vvv '16 January 2004
'20 January 2004 minor changes here
    II = InStr(1, vvv, "PHOTO ")
    temptemp = vvv
    If II <> 0 Then
        temptemp = Left(vvv, II - 1) + Right(vvv, Len(vvv) - (II + 6 - 1)) 'strip off "photo "
    End If
'
'16 August 2004    frmproj2.Caption = LCase(temptemp) + " (Replay by Spectate Swamp)" '17 January 2004
'    frmproj2.Caption = LCase(temptemp) + " (Replay by Spectate Swamp) length=" + CStr(video_length / 1000) + " seconds" '16 August 2004
    frmproj2.Caption = LCase(temptemp) + Trim(Cmd(79)) + " len=" + CStr(video_length / 1000) + " seconds" '11Jun2010 new element #79
'10 August 2010 over-ride the element 79 usage change this info when issuing a new copy to a golf course for their use etc
    frmproj2.Caption = LCase(temptemp) + " " + rand_str + " (By Spectate Swamp) len=" + CStr(video_length / 1000) + " secs" '16 August 2004
    If Left(Cmd(72), 11) = "RESULTS.TXT" Then                   '08 November 2004
        frmproj2.Caption = LCase(temptemp) + " (Logging to RESULTS.TXT Spectate Swamp) length=" + CStr(video_length / 1000) + " seconds" '16 August 2004
    End If          '08 November 2004
'    frmproj2.Caption = " Line_Search=" + Line_Search '08 November 2004 is the xxx. stuff for output
    If Left(UCase(Cmd(61)), 8) = "SETSPEED" Then
        If play_speed <> 1000 Then
'08 January 2004            delay_sec = delay_sec * ((1000 / play_speed) + 1)      '04 November 2003 if slower make dalay longer
            slomo = True        '08 January 2004 14 January 2004 moved 1 line up inside of the if ??
'            keep_slomo = True       '08Nov2011??
'            i = mci_setaudio(video1, 0, 0)  '29 February 2004
'           might need a combination of mci_setaudio and mcisendcommand todo ***vip***
        End If                              'add the 1 above just to extend the play a bit
'08 January 2004 set speed slow motion removed       i = mciSendString("set video1 speed " & play_speed, 0&, 0, 0)   '29 October 2003
    End If          '30 October 2003 switch on and off
'16 November 2003 set play_speed back to command file value
'14 January 2004 what is with the resetting here maybe too early..
'moved down    play_speed = save_play_speed        '16 November 2003
'
'08 November 2004 open access append then immediately close the file. here. results.txt open for output
'as well make a change to the heading above re the results.txt file logging info.
'    frmproj2.Caption = " TheFile=" + TheFile '08 November 2004 testing Just need "xxx." put in front here.
'    it sort of hangs if the logging is on and i am using it as input. this just makes sure. i did it first time
    If UCase(Left(Cmd(72), 11)) = "RESULTS.TXT" And UCase(Trim(TheFile)) <> "RESULTS.TXT" Then                   '08 November 2004
'    frmproj2.Caption = " ooo1=" + ooo '08 November 2004 testing Just need "xxx." put in front here.
       ResultFile = FreeFile        '08 November 2004
       Open "RESULTS.TXT" For Append Access Write As #ResultFile   '08 November 2004 do this if RESULTS.TXT IN Cmd(72)
        Print #ResultFile, LTrim(ccc)      '08 November 2004   this is the PHOTO stuff
        Print #ResultFile, Line_Search        '08 November 2004 this is the xxx. stuff
        Close ResultFile                            '08 November 2004
'then do a print / write then a close on that file
    End If                                                      '08 November 2004
'11 March 2007    If line_start_point <> 0 Then temp_double = line_start_point        '19 July 2003 Ver=1.07T
' testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secY=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info XX1 " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
    line_start_point = keep_line_start_point           '23Jan2012  somehow this is the fix......
    If line_start_point <> 10 Then temp_double = line_start_point        '19 July 2003 Ver=1.07T
'19 July 2003    If Left(UCase(Cmd(59)), 4) <> "WAIT" Or thumb_nail = "YES" Then
'18 August 2003    If thumb_nail = "YES" Or line_delay_sec > 0 Then
' testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secY=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info XX2 " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
    If delay_sec = 11111 Then
        delay_sec = video_length / 1000       '21Jan2012
        line_delay_sec = delay_sec
        begin = "YES"            '23Jan2012 changed from YES
        temp_double = 10       '23Jan2012
    End If                  '21Jan2012
    '23Jan2012
    If thumb_nail = "YES" Or InStr(aaa, "THUMB==") <> 0 Or line_delay_sec > 0 Then
'    If thumb_nail = "YES" Then
'        i = mciSendString("close video1", 0&, 0, 0)         '14 july 2003 moved from a few lines below
' testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secY=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info XX3 " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
      If temp_double > 1 Then
 '12 december 2004 test only           frmproj2.Caption = " Video to Start at " + CStr(temp_double) + " delay=" + CStr(line_delay_sec) '12 December 2004 testing
'testing to see what happens on the DVD with no read maybe a wait will fix it.????
'12 December 2004        i = mciSendString("play video1 from " + CStr(temp_double), 0&, 0, 0)  'this is where the display happens
' testtest = "begin=" + begin + " " + CStr(temp_double) + " " + CStr(begin_point) + " " + CStr(video_length) + " delay_secX=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
'       testtest = InputBox("check delay_sec info YY " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) '21Jan2012
    If begin = "YES" Then temp_double = 0       '15Sep2011
    If begin_point <> 0 Then temp_double = begin_point      '20Sep2011
' testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secX=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
' testtest = "begin=" + begin + " temp_double=" + CStr(temp_double)
        If Left(Cmd(75), 9) = "VIDEOSTOP" Then      '12 December 2004
'12Feb2012
'        temp_double = keep_line_start_point      '12Feb2012
' testing = "mpg_file11.5=" + mpg_file + " " + CStr(temp_double) + " " + Left(mssg, 10) + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + CStr(temp_double), "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'            DoEvents            '12feb2012
            i = mciSendString("play video1 from " + CStr(temp_double), 0&, 0, 0) 'this is where the display happens
'23Jan2013 try here
    If debug_photo Then         '23Jan2013
'Print "testing 24Jan2013 only"
'        i = mciSendString("pause video1", 0&, 0, 0)  '24Jan2013
'            i = mciSendString("stop video1", 0&, 0, 0) '24Jan2013
'            i = mciSendString("close video1", 0&, 0, 0) '24Jan2013
'            SetFocus '24Jan2013
'Print "testing 24Jan2013 only"
            tt1 = InputBox("testing photo 7b " + Pict_file, , , 4400, 4500) 'TESTING ONLY
    End If
'            i = mciSendString("pause video1", 0&, 0, 0) '24Jan2013
'            Print "testing 24Jan2013"
'            DoEvents            '12feb2012
'            DoEvents            '12feb2012
'            DoEvents            '12feb2012
'            DoEvents            '12feb2012
'            DoEvents            '12feb2012
'            i = mciSendString("status video1 mode", mssg, 255, 0) '12feb2012 see if another call fixes it
'12Feb2012
'        temp_double = keep_line_start_point      '12Feb2012
' testing = "mpg_file11.5=" + mpg_file + " " + CStr(temp_double) + " " + Left(mssg, 10) + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'10feb2012 see what is happening here  with this set below the video runs complete
' testing = "mpg_file15=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            Else
        End If                                      '12 December 2004
      Else
'testtest = CStr(temp_double) + " " + CStr(line_delay_sec) + " " + CStr(video_length) + " delay_secX1=" + CStr(delay_sec) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '21Jan2012 teststring
' testtest = "begin=" + begin + " temp_double=" + CStr(temp_double)
            i = mciSendString("play video1", 0&, 0, 0)   'this is where the display happens  play video ***vip***
      End If
'        xtemp = InputBox(" 24 March 2004 nowait test", "testing Prompt   " + CStr(line_start_point) + " " + CStr(delay_sec) + " " + CStr(line_delay_sec), , xx1 - offset1, yy1 - offset2) '14 july 2003 test
      
        GoSub line_30000                    'do a delay test
    Else
'        xtemp = InputBox(" nowait test", "testing Prompt   ", , xx1 - offset1, yy1 - offset2)  '14 july 2003 test
'18 March 2004 need to trap any errors on the start of videos here
' i = mciSendString("close all wait", 0&, 0, 0)      '17 March 2004 make sure nothing is there..
' i = mciSendString(todo$, 0&, 0, 0)
    On Error GoTo trap2         '18 March 2004
    DoEvents                    '18 March 2004
      If temp_double > 1 Then
'16 March 2004        i = mciSendString("play video1 from " + CStr(temp_double) + " wait", 0&, 0, 0) 'this is where the display happens (noWAIT)
'16 March 2004 This is where the video plays displays shows etc ***vip***
' testing = "mpg_file1.1=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
    i = mciSendString("play video1 from " + CStr(temp_double), 0&, 0, 0)  'this is where the display happens (noWAIT)
'    i = mciSendString("play video1 from " + CStr(temp_double) + " wait", 0&, 0, 0) 'this is where the display happens (noWAIT)
       End If               '18 March 2004
'14 March 2004 I think the wait change above was a major factor in getting the video playback not to bomb out
'14 March 2004        i = mciSendString("play video1 from " + CStr(temp_double), 0&, 0, 0)  '08 March 2004
'18 March 2004      Else
'09 september 2003 windows 2000 service pack 1 / windows xp has problems here
'        xtemp = InputBox(" test 6bb", "testing Prompt   ", , xx1 - offset1, yy1 - offset2)  '09 september 2003 test

'***the following logic never seems to get executed the start is always 10 or greater
     If temp_double <= 1 Then
'        i = mciSendString("play video1 wait", 0&, 0, 0) 'this is where the display happens (noWAIT)
' testing = "mpg_file1.2=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
        i = mciSendString("play video1", 0&, 0, 0) 'this is where the display happens (noWAIT)
'        xtemp = InputBox(" test 6bbb", "testing Prompt   ", , xx1 - offset1, yy1 - offset2)  '09 september 2003 test
      End If
    GoTo past_trap2             '17 March 2004
trap2:                          '17 March 2004
    Resume past_trap2
        xtemp = InputBox("trap2 error", CStr(video_length), , xx1 - offset1, yy1 - offset2)  '17 March 2004
            new_delay_sec = 5.5
            GoSub line_30300        '18 March 2004
past_trap2:     '18 March 2004
    DoEvents                    '18 March 2004
    On Error GoTo Errors_31000  '18 March 2004
    DoEvents                    '18 March 2004
'18 March 2004
      
    End If                  '14 July 2003
    play_speed = save_play_speed        '16 November 2003 14 January 2004 moved down to here
'24 September 2003
'    GoSub line_30000        '24 September 2003 test this
    '24 September 2003 freeze
    
    '26 February 2004 for the freeze of a video do not use the line_30000 routine???
    '   because we are doing stops and starts that surely won"t work
    
        i = mciSendString("pause video1", 0&, 0, 0)  '14 March 2004 video playing displaying at this point ***vip***
'10Feb2012 testing only below
' testing = "video_audit3=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
        DoEvents                                        '14 March 2004
    If Line_freeze_sec <> 0 Then
        new_delay_sec = Line_freeze_sec
    'need a mcisendstring that will do a stop eg below
        DoEvents        '05 March 2004
'29 February 2004
'17 March 2004 activate the lines below re the position
'19 March 2004 get the position anyway
'20 March 2004
    If slomo = False Then                          '20 March 2004
'20 March 2004        If Len(Trim(tempdata)) < 1 Then
            i = mciSendString("status video1 position", vs, 255, 0) '29 February 2004
            temp3 = InStr(vs, Chr$(0)) '19 March 2004
            temp11 = Val(Left(vs, temp3 - 1)) '19 March 2004
            tempdata = CStr(temp11)     '19 March 2004
    End If      '19 March 2004
        DoEvents            '09 March 2004
'14 March 2004        i = mciSendString("stop video1", 0&, 0, 0)  '09 March 2004
'        i = mciSendString("pause video1", 0&, 0, 0)  '05 March 2004
        DoEvents        '06 March 2004
        If Len(tempdata) < 1 Then
            frmproj2.Caption = LCase(temptemp) + "(Video length=" + CStr(video_length) + ")" '29 Fe
            Else
            frmproj2.Caption = LCase(temptemp) + "(Freeze at " + tempdata + " of " + CStr(video_length) + ")" '29 February 2004
        End If              '17 March 2004
            tempdata = ""   '18 March 2004
'17 March 2004            frmproj2.Caption = LCase(temptemp) + "(Freeze)" '15 March 2004
'26 February 2004        GoSub line_30000            '24 September 2003
        GoSub line_30300            '26 February
        DoEvents            '09 March 2004
'dougdoughere maybe comment out below
'11 March 2004        i = mciSendString("resume video1", 0&, 0, 0)  '09 March 2004
        DoEvents            '09 March 2004
        GoTo line_17053         '01 October 2003
    End If             '24 September 2003
'01 October 2003
    If freeze_sec <> 0 Then
        new_delay_sec = freeze_sec
    'need a mcisendstring that will do a stop eg below
        DoEvents            '09 March 2004
'15 March 2004 somehow by removing the position call below seems to eliminate a major crash
'17 March 2004  activate the position stuff below
'17 March 2004 temp            i = mciSendString("status video1 position", vs, 255, 0) '29 February 2004
'17 March 2004 temp            tempdata = CStr(Val(vs))     '26 February 2004
'20 March 2004
    If slomo = False Then                          '20 March 2004
'20 March 2004        If Len(Trim(tempdata)) < 1 Then
            i = mciSendString("status video1 position", vs, 255, 0) '29 February 2004
            temp3 = InStr(vs, Chr$(0)) '19 March 2004
            temp11 = Val(Left(vs, temp3 - 1)) '19 March 2004
            tempdata = CStr(temp11)     '19 March 2004
    End If      '19 March 2004
        DoEvents        '06 March 2004
'14 March 2004        i = mciSendString("stop video1", 0&, 0, 0)  '09 March 2004
'        i = mciSendString("pause video1", 0&, 0, 0)  '05 March 2004
        If tempdata = "" Then
            frmproj2.Caption = LCase(temptemp) + "(Video length=" + CStr(video_length) + ")" '29 Fe
            Else
'28Jan2012            frmproj2.Caption = LCase(temptemp) + "(Freeze at " + tempdata + " of " + CStr(video_length) + ")" '29 February 2004
'31Jan2012 let cmd(60) control this            If pauseyn = "Y" Then new_delay_sec = 0        '28Jan2012
            frmproj2.Caption = LCase(temptemp) + " secs== " + CStr(freeze_sec) + " (Freeze at " + tempdata + " of " + CStr(video_length) + ")" '29 February 2004
        End If              '17 March 2004
'            frmproj2.Caption = LCase(temptemp) + "(Freeze)" '15 March 2004
'26 February 2004        GoSub line_30000            '24 September 2003
        GoSub line_30300            '26 February
'        xtemp = InputBox(" test freeze=  " + Format(Line_freeze_sec, "###0.000"), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
        DoEvents            '09 March 2004
'        GoTo line_17053         '01 October 2003
'dougdoughere       maybe comment out below
'11 March 2004        i = mciSendString("resume video1", 0&, 0, 0)  '09 March 2004
        DoEvents            '09 March 2004
    End If             '01 October 2003

line_17053:             '01 October 2003
        i = mciSendString("stop all", 0&, 0, 0)  '14 March 2004
'10Feb2012 testing only below
' testing = "video_audit4=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
        DoEvents                                        '14 March 2004

'01 March 2004 check status here before continuing...
'            i = mciSendString("status video1 position", vs, 255, 0) '01 March 2004
'            i = mciSendString("status video1 length", mssg, 255, 0) '01 March 2004
'06 March 2004 check out the ready status here first..
            DoEvents        '06 March 2004
            i = mciSendString("status video1 ready", temps, 12, 0) '06 March 2004
'10Feb2012 testing only below
' testing = "video_audit5=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            DoEvents        '06 March 2004
            i = mciSendString("status video1 mode", mssg, 255, 0) '01 March 2004
'10Feb2012 testing only below
' testing = "video_audit6=" + mpg_file + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            DoEvents        '06 March 2004
            frmproj2.Caption = "hey 2 " + Left(temps, 4) + " " + Left(mssg, 5) '09 March 2004
            DoEvents        '06 March 2004
'              End If
'dougheredoughere
        If Left(UCase(mssg), 7) = "STOPPED" Then
            DoEvents                                    '05 March 2004
            frmproj2.Caption = "hey 2a " + Left(temps, 4) + " " + Left(mssg, 5) '09 March 2004
            DoEvents        '06 March 2004
            GoTo line_17054                     '11 March 2004
       End If              '05 March 2004
        
        If Left(UCase(mssg), 7) = "PLAYING" Then
            DoEvents                                    '05 March 2004
'dougheredoughere
'11 March 2003            i = mciSendString("pause video1", 0&, 0, 0) '05 March 2004
'            i = mciSendString("stop video1", 0&, 0, 0) '09 March 2004
            DoEvents                                    '05 March 2004
'            i = mciSendString("close video1", 0&, 0, 0)         '14 july 2003 moved from a few lines below
            DoEvents                            '11 March 2004
            frmproj2.Caption = "hey 3 " + Left(temps, 4) + " " + Left(mssg, 5) '11 March 2004
            DoEvents                            '11 March 2004
            GoTo line_17054                     '11 March 2004
        End If              '02 March 2004 added the if statement around the mcisendstring above
        
        If Left(UCase(mssg), 6) = "PAUSED" Then
'            i = mciSendString("close video1", 0&, 0, 0)         '06 March 2004
            frmproj2.Caption = "hey 4 " + Left(temps, 4) + " " + Left(mssg, 5) '11 March 2004
            DoEvents                            '11 March 2004
            GoTo line_17054                     '11 March 2004
        End If              '06 March 2004
        
            frmproj2.Caption = "hey 5 " + Left(mssg, 5) + " " + Left(temps, 5) '11 March 2004
            GoTo line_17054a    '11 March 2004
line_17054:                 '11 March 2004
            frmproj2.Caption = "hey 6 " + Left(mssg, 5) + " " + Left(temps, 5) '11 March 2004
line_17054a:                '11 March 2004
'dougdoughere uncomment out below
'            i = mciSendString("close video1", 0&, 0, 0)         '09 March 2004
'28Jan2012 might want to skip the below statement if pauseyn = "Y"
    If pauseyn = "Y" Then GoTo no_close             '28Jan2012
            i = mciSendString("close video1", 0&, 0, 0)         '05 April 2004
'16 March 2004            i = mciSendString("close all wait", 0&, 0, 0)         '14 March 2004
no_close:               '28Jan2012
           pauseyn = hold_pauseyn       '25Mar2012
        On Error GoTo close_error
next_close_error:            '16 March 2004
            DoEvents
'            new_delay_sec = 0.1     '18 March 2004
'            GoSub line_30300        '18 March 2004 maybe the delay will fix the problem
            'do a pause here         18 March 2004
'04 April 2004            i = mciSendString("close all", 0&, 0, 0)         '14 March 2004
            frmproj2.Caption = "hey 6a " + Left(mssg, 5) + " " + Left(temps, 5) '11 March 2004
'            new_delay_sec = 0.1     '18 March 2004
'            GoSub line_30300        '18 March 2004 maybe the delay will fix the problem
            'do a pause here         18 March 2004
            DoEvents                '11 March 2004
        On Error GoTo Errors_31000        '16 March 2004
'           Set colReminderPages = Nothing  'release memory?? 16 March 2004
        GoTo past_close_error           '16 March 2004
close_error:                            '16 March 2004
'        On Error GoTo Errors_31000      '16 March 2004
'        i = mciSendString("status video1 mode", mssg, 255, 0) '16 March 2004
        Resume close_error_1       '16 March 2004   clear the error raised by the close
close_error_1:
        DoEvents                    '16 March 2004
            new_delay_sec = 5.5
            GoSub line_30300        '16 March 2004 maybe the delay will fix the problem
        frmproj2.Caption = LCase(temptemp) + Left(mssg, 4) + " (Close Error)" + CStr(line_start_point + (line_delay_sec * 1000)) + " > " + CStr(vs) '26 February 2004
            new_delay_sec = 5.5
            GoSub line_30300        '16 March 2004 maybe the delay will fix the problem
            'do a pause here
            GoTo next_close_error           '16 March 2004
past_close_error:
'            new_delay_sec = 0.1     '16 March 2004 it may still be over-running here and maybe not
'                                    '16 March 2004 may want to test without some time.
'            GoSub line_30300        '16 March 2004 maybe the delay will fix the problem
            DoEvents        '11 March 2004
            frmproj2.Caption = "hey 7 " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '11 March 2004
'            old_line = temptemp + ""        '31 March 2004
            On Error GoTo Errors_31000      '18 March 2004 just to make sure it is set back...
'           xtemp = InputBox(" 05 March 2004 test  ", "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
'14 july 2003 i = mciSendString("close video1", 0&, 0, 0)
    If elapse_yn = "YES" Then
        elapse_end = Timer            '13 July 2003
        xtemp = InputBox(" elapse=  " + Format(elapse_end - elapse_start, "###0.000"), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)

    End If                            '13 July 2003
    slomo = False                       '08 January 2004
    slomo = True                '31Dec2011
  End If
    GoTo Line_17055
End If                      '11 May 2003 ver=1.05
        
'   **** this is a very long if statement ****  **hint**
        'end of if statement "If mpg_file = "YES" Then"
        
'10 June 2003            wav file display here ver=1.07
If wav_file = "YES" Then
    motion_yn = "YES"
    If videoyn = "SHOWVIDEO" Then
Last$ = frmproj2.hWnd & " Style " & &H40000000
' ToDo$ = "open e:\cottonwood\101-0186_mvi.avi Type avivideo Alias video1 parent " & Last$
 todo$ = "open " + Pict_file + " Type waveaudio Alias video1 parent " & Last$
 i = mciSendString(todo$, 0&, 0, 0)
 i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
 i = mciSendString("play video1 wait", 0&, 0, 0)
 i = mciSendString("close video1", 0&, 0, 0)
    End If
    GoTo Line_17055
End If                      '10 June 2003 ver=1.05
     
'10 June 2003            midi file display here ver=1.07
If mid_file = "YES" Then
    motion_yn = "YES"
    If videoyn = "SHOWVIDEO" Then
Last$ = frmproj2.hWnd & " Style " & &H40000000
' ToDo$ = "open e:\cottonwood\101-0186_mvi.avi Type avivideo Alias video1 parent " & Last$
 todo$ = "open " + Pict_file + " Type sequencer Alias video1 parent " & Last$
 i = mciSendString(todo$, 0&, 0, 0)
 i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
 i = mciSendString("play video1 wait", 0&, 0, 0)
 i = mciSendString("close video1", 0&, 0, 0)
    End If
    GoTo Line_17055
End If                      '10 June 2003 ver=1.05
     
     
     
'    If Cmd(56) = "PHOTO_DETAIL" Then        '22 March 2003 ver=1.02
'        Cls
'    End If                                  '22 March 2003 ver=1.02 just clear screen
    
'03 August 2003
'        tt1 = InputBox("doug testing " + aaa, , , 4400, 4500) 'TESTING ONLY 03 August 2003
    line_fit = ""                '03 August 2003
    If InStr(1, UCase(aaa), "FIT==") <> 0 Then
            line_fit = "FIT"
    End If                  '03 August 2003
    
    If InStr(1, UCase(aaa), "REG==") <> 0 Then
            line_fit = "REG"
    End If                  '03 August 2003
    
'03 August 2003    If img_ctrl = "YES" Then
'        Set Image1.Picture = LoadPicture(Pict_file) 'Stretch Mode
'    Else
'        Set Picture = LoadPicture(Pict_file)        'Normal Mode
'    End If              'february 21 2001
'        xtemp = InputBox(" testing doug#77  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)

    If (img_ctrl = "YES" Or line_fit = "FIT") And line_fit <> "REG" Then
'===========================================================
    '28 November 2004 this trap fixed the problem with the bad .gif files
    '   but it does not trap the problem associated with the "malletc031.jpg" file that I was having.
    '   the sequence would not go to the text display when I was using the control== sequence somehow I will keep the file for later
        On Error GoTo first_problem  '28 November 2004
            DoEvents                    'march 18 2001
        Set Image1.Picture = LoadPicture(Pict_file) 'Stretch Mode
            On Error GoTo Errors_31000      '28 November 2004
            GoTo frst_17055
first_problem:       '28 November 2004
            frmproj2.Caption = "bad file= " + Pict_file '28 November 2004
'        xtemp = InputBox("file display problem=" + CStr(Err.Number) + " " + Err.Description, , , xx1 - offset1, yy1 - offset2) 'march 18 2001
        On Error GoTo Errors_31000    '28 November 2004
        Resume input_1000      '28 November 2004
frst_17055:                      '28 November 2004
'===========================================================
'28 November 2004 original was here        Set Image1.Picture = LoadPicture(Pict_file) 'Stretch Mode
        line_fit = "FIT"                '08 August 2003
'        xtemp = InputBox(" testing doug#6  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

    If (img_ctrl <> "YES" And line_fit <> "FIT") Or line_fit = "REG" Then
'    If img_ctrl <> "YES" And line_fit <> "FIT" Then
'    save_line = "17055xxy"            '28 November 2004 testing for bug
'===========================================================
        On Error GoTo display_problem  '28 November 2004
            DoEvents                    'march 18 2001
        Set Picture = LoadPicture(Pict_file)        'Normal Mode
            On Error GoTo Errors_31000      '28 November 2004
            GoTo bef_17055
display_problem:       '28 November 2004
            frmproj2.Caption = "bad file= " + Pict_file '28 November 2004
'        xtemp = InputBox("file display problem=" + CStr(Err.Number) + " " + Err.Description, , , xx1 - offset1, yy1 - offset2) 'march 18 2001
        On Error GoTo Errors_31000    '28 November 2004
        Resume input_1000      '28 November 2004
bef_17055:                      '28 November 2004
'===========================================================
'original was here 28 November 2004        Set Picture = LoadPicture(Pict_file)        'Normal Mode
'        xtemp = InputBox(" testing doug#7  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001

Line_17055:                 '01 february 2003
        save_line = "17055"

'10Feb2012 testing only below
'            DoEvents
'            i = mciSendString("status video1 ready", mssg, 12, 0)
' testing = "video_audit7=" + mssg + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'            DoEvents
'            i = mciSendString("status video1 mode", mssg, 255, 0)
'  testing = "video_audit8=" + mssg + " line_freeze_sec=" + CStr(Line_freeze_sec) + " delay_sec=" + CStr(delay_sec) + " begin=" + begin + " line_start_point=" + CStr(line_start_point) + " video_length=" + CStr(video_length) + "begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) '17Jan2012 teststring
'       testing = InputBox("check delay_sec info " + testing, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'           DoEvents
'10feb2012 right after this it should run testing above lines
'            frmproj2.Caption = "hey 7a " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004

'20 march 2003 ver=1.02 testing of call to last_lines_15000
'
    
'   See the Logic for post photo text ie day and time at line_2130
'        xtemp = InputBox(" 08 August 2003 test  " + sscreen_saver + " " + Left(Cmd(56), 12), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)

'18 November 2004    If UCase(Left(Cmd(56), 12)) = "PHOTO_DETAIL" Then       '22 March 2003 ver=1.02
    If detailyn = "PHOTO_DETAIL" Then       '18 November 2004
'21 March 2004        If sscreen_saver = "Y" Then GoSub line_30000    'do a extra pause again on "weekday dd Month" display

            new_delay_sec = Val(Cmd(27))    '03 September 2004
        If sscreen_saver = "Y" And motion_yn <> "YES" Then GoSub line_30300    '21 March 2004 do not do a pause if video
'03 September 2004        If sscreen_saver = "Y" And mpg_file <> "YES" Then GoSub line_30300    '21 March 2004 do not do a pause if video
'03 September 2004        If sscreen_saver = "Y" And mpg_file <> "YES" Then GoSub line_30000    '21 March 2004 do not do a pause if video
        If line_fit = "FIT" Then            '08 August 2003 testing only
            Set Image1.Picture = LoadPicture() '08 August 2003 testing only
            If Not mixx Then Set Picture = LoadPicture(Pict_file)           '12Feb2017 maybe
'            Set Picture = LoadPicture(Pict_file)        'Normal Mode 08 august 2003
        End If                              '08 August 2003 testing only
        
        If stretch_image = "YES" Then
            line_fit = "FIT"
        Else
            line_fit = "REG"
        End If              '08 august 2003

'        GoSub line_30000        'pause for the wait time before text display....
                            'will want to make this optional... or cut the time in half??

    '   maybe resetting the stuff below will allow for hi-liting in the displayed characters
    '   Not needed the values in SSS1 etc were fine?
'        SSS1 = SAVE_KEEPS1
'        SSS2 = SAVE_KEEPS2
'        SSS3 = SAVE_KEEPS3
'        SSS4 = SAVE_KEEPS4
'        SSS5 = SAVE_KEEPS5
'        SSS6 = SAVE_KEEPS6
    
'        SSS1 = KEEPS1
'        SSS2 = KEEPS2
'        SSS3 = KEEPS3
'        SSS4 = KEEPS4
'        SSS5 = KEEPS5
'        SSS6 = KEEPS6
    
'            tt1 = InputBox("testing SSS1...=" + SSS1 + SSS2 + SSS3 + SSS4 + SSS5 + SSS6, , , 4400, 4500) 'TESTING ONLY
    lll = old_line          'use this one in stead non capitalized ???
'31 March 2004    ooo = lll               '*** make this optional later *** and probably remove "photo " now
     
     '31 March 2004
                'only problem so far is that lll is all uppercase (create and use old_line) instead
'31 March 2004 ooo should be the original data.
    array_pos = 0       '31 March 2004 this one caused a bit of grief
    ooo = lll   '31 March 2004
    aaa = UCase(ooo) + "" 'seems aaa needs to have info for prints
    
'31 March 2004 if a double quote " found and an end quote with it use just that for display
    
    II = InStr(1, ooo, Chr$(34))   '31 March 2004 check for info contained in quotes
    If II <> 0 Then
        tt = InStr(II + 1, ooo, Chr$(34))
        If tt <> 0 Then
            ooo = Mid(ooo, II + 1, tt - II - 1)
            GoTo line_17070
        End If
    End If
'31 March 2004 above logic added
    
    
    
    II = InStr(1, UCase(ooo), "PHOTO ")
    If II <> 0 Then
        ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - (II + 6 - 1)) 'strip off "photo "
    End If
'29 March 2004
    II = InStr(1, UCase(ooo), "START==")
    If II <> 0 Then
        tt = InStr(II + 7, ooo, " ")
        If tt <> 0 Then
            ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - tt)
        End If
    End If      '29 March 2004
    
'29 March 2004
    II = InStr(1, UCase(ooo), "SPEED=")
    If II <> 0 Then
        tt = InStr(II + 6, ooo, " ")
        If tt <> 0 Then
            ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - tt)
        End If
    End If      '29 March 2004
    
'29 March 2004
    II = InStr(1, UCase(ooo), "WAIT=")
    If II <> 0 Then
        tt = InStr(II + 5, ooo, " ")
        If tt <> 0 Then
            ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - tt)
        End If
    End If      '29 March 2004
    
'29 March 2004
    II = InStr(1, UCase(ooo), "FREEZE=")
    If II <> 0 Then
        tt = InStr(II + 7, ooo, " ")
        If tt <> 0 Then
            ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - tt)
        End If
    End If      '29 March 2004
    
'29 March 2004
    II = InStr(1, UCase(ooo), "LENGTH=")
    If II <> 0 Then
        tt = InStr(II + 7, ooo, " ")
        If tt <> 0 Then
            ooo = Left(ooo, II - 1) + Right(ooo, Len(ooo) - tt)
        End If
    End If      '29 March 2004
    
    
    
line_17070:             '31 March 2004
    
'    aaa = "dummydummydummydummydummydummydummydummy"
    'over-riding the statement below makes for a better display for now...
    'should test with the following active just so it can be used ie hi-liting searched string
    'text where the other when no matches found shows whole line in one color. that is all
    'it works fine now the logic at has been deactivated.... at line_12600 "screen_saver = "N""
    'maybe works better than fine.....
    aaa = UCase(ooo) '+ "============================================"          'last change here 26 March 2003 10:00 am
    If mixx Then GoTo line_17075           '12Feb2017
'    ooo = ccc + ""          'to display full line *** make this optional later ***
    Font.Size = 48
    Font.Bold = True
'            frmproj2.Caption = "hey 7b " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004

'    Font.Italic = True 'jff just for fun
'    SetFocus
'    Set Picture = LoadPicture("c:\search\Sparrow.jpg")
 '   frmproj2.Caption = program_info + " (cls #9)" '18Dec2013
    Cls                 'just to clear anything
'    Def_Fore = 14       'make the text yellow
'    Def_Fore = 10       'make the text lime green instead of yellow???
'20 May 2003 (use what is in control file)    Def_Fore = 13       'make the text light purple pink???

'    Def_Fore = 7        'try grey
    ForeColor = QBColor(Def_Fore)
    Context_lines = 7   '28 March 2003 change from 7 to 8
    line_len = 30       'have it wrap after 20 characters
    'a lot of messing around just to set above to 60 for now on display
    
'30 March 2004    Clear_Context_lines = Context_lines
'30 March 2004    Context_cnt = Context_lines 'need a dummy context fields set up here
'30 March 2004    For II = 1 To Context_cnt
'30 March 2004    For II = 1 To 40        '30 March 2004
'30 March 2004        Context_text(II) = Space(20)
'30 March 2004    Next II
    
    '     endstuff testing
    '        endstuff = "NO"     '27 March 2003 test getting it all out with this...
            
'            tt1 = InputBox("testing array_ooo(array_prt)=" + endstuff + CStr(array_pos) + CStr(array_prt) + array_ooo(array_prt), , , 4400, 4500) 'TESTING ONLY

'        If array_pos <> 0 And array_prt > 0 Then
'            tt1 = InputBox("testing array_ooo(array_prt)=" + endstuff + CStr(array_pos) + CStr(array_prt) + array_ooo(array_prt), , , 4400, 4500) 'TESTING ONLY
'        End If

'            tt1 = InputBox("testing ooo=" + endstuff + CStr(array_pos) + CStr(array_prt) + ooo, , , 4400, 4500) 'TESTING ONLY
'            tt1 = InputBox("testing endstuff data_ooo=" + endstuff + data_ooo, , , 4400, 4500) 'TESTING ONLY
'            tt1 = InputBox("testing SSS1=" + SSS1 + SSS2 + SSS3 + SSS4, , , 4400, 4500) 'TESTING ONLY
'            tt1 = InputBox("testing keeps1=" + KEEPS1 + KEEPS2 + KEEPS3 + KEEPS4, , , 4400, 4500) 'TESTING ONLY
'the following lines of info should print right on top of the photo but
'somehow it is not. see the smg emulation where a bmp file is displayed as a pict
'then the text is written and re-written over that image??? langaliers where are you
'*** note seems that the text can only be displayed if "normal" mode ie see normal above
'    the "Set Image1.Picture = LoadPicture(Pict_file) from above won't do it... ie image control..

'was code here related to 1.02b but was moved
'29 March 2003 test change for ver=1.04 below
    '03 August 2003
'03 August 2003    If img_ctrl = "YES" And motion_yn <> "YES" Then
'    If (img_ctrl = "YES" Or line_fit = "FIT") And motion_yn <> "YES" Then
'    If img_ctrl = "YES" And motion_yn <> "YES" And line_fit <> "REG" Then
'03 August 2003    If img_ctrl = "YES" And motion_yn <> "YES" And line_fit <> "FIT" Then
'03 August 2003 test 1    If img_ctrl <> "xxxx" And motion_yn <> "YES" Then          '03 August 2003 just try and override
    If img_ctrl = "xxxx" And motion_yn <> "YES" Then          '03 August 2003 just try and override
'        xtemp = InputBox(" testing doug#8-  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
        Set Image1.Picture = LoadPicture() 'Stretch Mode clear away here
            If Not mixx Then Set Picture = LoadPicture(Pict_file)           '12Feb2017
'        Set Picture = LoadPicture(Pict_file)        'Normal Mode
'dougheredoughere
'03 August 2003 remove this stay with above        Set Picture = LoadPicture()        'Normal Mode lp#6
'        xtemp = InputBox(" testing doug#8  " + aaa + "*line_fit=" + line_fit + "*img_ctrl=" + img_ctrl, "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    End If              'february 21 2001
'03 August 2003 test 1
    If img_ctrl = "YES" And motion_yn <> "YES" And line_fit = "REG" Then
        Set Picture = LoadPicture(Pict_file)        'normal mode
    End If          '03 August 2003 test 1
    
line_17075:         '12Feb2017
    If mixx Then GoTo Line_17090           '12Feb2017
'12Feb2017  need above
    If img_ctrl <> "YES" And line_fit = "FIT" Then
        Set Image1.Picture = LoadPicture() 'Stretch Mode clear away here
    End If          '03 August 2003 test 1
'        Set Image1.Picture = LoadPicture() '08 August 2003 testing
'30 March 2004    GoSub Last_lines_15000
'    frmproj2.Caption = program_info + " (cls #10)" '18Dec2013
    Cls                 '30 March 2004
    Print               '30 March 2004
    Print               '30 March 2004
    Print               '30 March 2004
    Print               '30 March 2004
    Print               '30 March 2004
    Print       '30 March 2004
    Print       '30 March 2004
'            tt1 = InputBox("30 March 2004 " + CStr(Context_lines), , , 4400, 4500) 'TESTING ONLY
    GoSub sub_12000     'do the bolding and high-liting hi-liting here
'            tt1 = InputBox("31 March 2004 " + Left(ooo, 20), , , 4400, 4500) 'TESTING ONLY
    Font.Size = Val(Cmd(2))     'reset it back from 48 above
'13 April 2004    new_delay_sec = Val(Cmd(27))  '30 March 2004
    new_delay_sec = Val(Cmd(60))  '13 April 2004
'30 March 2004    If mpg_file = "YES" Then GoSub line_30300    '30 March 2004
    GoSub line_30300    '30 March 2004

line_17075a:         '12Feb2017
    End If      'end of PHOTO_DETAIL check pretty well ver=1.02 to here
    '12Feb2017 need to skip the above if statement when doing a mix text / picture (do not want font change etc)
'            frmproj2.Caption = "hey 7c " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004
    

'    Print "what the hey"        '21 march 2003
'            tt1 = InputBox("testing b=" + endstuff + "=" + vvv + " tot_print=" + CStr(tot_print), , , 4400, 4500) 'TESTING ONLY
'            tt1 = InputBox("testing b=" + CStr(Clear_Context_lines) + "=", , , 4400, 4500) 'TESTING ONLY
'20 march 2003 ver=1.02 ^
        
        last_pict = Pict_file   '28 february 2003
        
        dsp_cnt = dsp_cnt + 1           'may 12 2001
        previous_count = previous_count + 1   'october 9 2000
        If previous_count > 100 Then
            previous_count = 1
        End If
'        previous_picture(previous_count) = 0
        previous_picture(previous_count) = zzz_cnt
              'save line number of previous picture display
'        Print "b previous_picture(previous_count)zzz_cnt,previous_count"; previous_picture(previous_count); "="; zzz_cnt, previous_count
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
'
'        MyAppID = Shell(Cmd(8), 3)
'        AppActivate MyAppID
'here try and do a pause or wait as follows
'           For temp1 = 1 To 100000000
'           DoEvents
'           Next temp1
'           SendKeys "^o", True
'           SendKeys Pict_file, True
'           SendKeys "~", True    'FORCE AN "ENTER"
    
    End If
    'endiffy 17050
'            frmproj2.Caption = "hey 7d " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004
    
    save_line = "17055a"
      If InStr(UCase(Cmd(8)), "EXPLORER") <> 0 And Test1_str = "P2" Then
         MyAppID = Shell(Cmd(8) + " / select, " + Pict_file, 3)
         SendKeys "^o", True
      End If
    save_line = "17055b"
       If InStr(UCase(Cmd(8)), "EXPLORER") <> 0 And Test1_str = "P3" Then
 
 '          MyAppID = Shell(Cmd(8) + " / select, " + Pict_file, 3)
  '      Print "Cmd(8)"; Cmd(8); " Pict_file="; Pict_file
   '     tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
           
           MyAppID = Shell("command.com /c " + Cmd(8) + " / select, " + Pict_file, 3)
'         AppActivate MyAppID
'           SendKeys Pict_file, True
'           SendKeys "^v", True
       End If
    save_line = "17055c"
      If InStr(UCase(Cmd(8)), "EXPLORER") <> 0 And Test1_str = "P" Then
         MyAppID = Shell(Cmd(8) + " / select, " + Pict_file, 3)
         SendKeys "^o", True
      End If
        'here testing below netscape.exe
'the P1 option should work for .bmp .jpg and .tif
    save_line = "17055d"
    If InStr(UCase(Cmd(8)), "NETSCAPE") <> 0 And Test1_str = "P1" Then

   
            If Not mixx Then Set Picture = LoadPicture(Pict_file)           '12Feb2017
'        Set Picture = LoadPicture(Pict_file)        'normal
    End If
    save_line = "17055e"
'            Print "Cmd(8)="; Cmd(8); "="; Test1_str
 '       tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY

      If InStr(UCase(Cmd(8)), "NETSCAPE") <> 0 And Test1_str = "P" Then
        MyAppID = Shell(Cmd(8) + " " + Pict_file, 3)
        SendKeys "^v", True
      End If
Line_17080:
 '       AppActivate MyAppID
'            frmproj2.Caption = "hey 7e " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '01 September 2004
      
'   SendKeys "^F", True
'    SendKeys Pict_file, True
 '   SendKeys LastFile, True
 '   SendKeys "", True
'added the Test1_str check below aug 08/00
    If InStr(UCase(Line_Search), ".BMP") <> 0 And Test1_str = "P" Then
'    If InStr(UCase(Line_Search), ".BMP") <> 0 Then
    'check the program exists before going anywhere
    save_line = "17080"
    FileFile = FreeFile
    Close FileFile
    Open Cmd(9) For Input As #FileFile
    Close FileFile
'        Print "file name="; Pict_file; "="; Test1_str
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY
         
        MyAppID = Shell(Cmd(9), 1)
        SendKeys "^o", True
        SendKeys Pict_file, True
        SendKeys "~", True    'FORCE AN "ENTER"
    End If
    save_line = "17059b"
'            Set Picture = LoadPicture() 'clear any picture
    If InStr(UCase(Line_Search), ".TIF") <> 0 Then
    'check the program exists before going anywhere
    save_line = "17085"
    FileFile = FreeFile
    Close FileFile
    DoEvents
    Open Cmd(16) For Input As #FileFile
    Close FileFile
 
        LastFile = Pict_file
        MyAppID = Shell(Cmd(16) + " / select, " + Pict_file, 1)

    SendKeys "^o", True
    SendKeys "^v", True
    End If
    

         
Line_17090:
'    TheFile = ""
            frmproj2.Caption = "hey 7f " + Left(mssg, 5) + " " + CStr(array_pos) ' + Left(temptemp, 15) '21Mar2016
Return          'end of display_pict_17000

line_17200:     'February 24 2002 do the append of the file here
    
    'skip the file if the output file is the same as the one to append....
    If InStr(UCase(Test1_str), UCase(FileExt)) <> 0 Then
'        tt1 = InputBox("testing merge logic " & Test1_str, , , 4400, 4500) 'TESTING ONLY
'        If tt1 = "X" Or tt1 = "x" Then
'            GoTo End_32000
'        End If              'testing only
'            Print "doug testing "; Test1_str, FileExt
            GoTo line_17230
    End If
    f = f + 1
    Print "File merging "; Test1_str, FileExt
    DoEvents
    If f > 20 Then
        f = 1
        Cls
    End If
    save_line = "17200"
            Open Test1_str For Input As FileFile
line_17210:
            Line Input #FileFile, aaa
                Print #ExtFile, aaa
            GoTo line_17210
line_17220:
        Print #ExtFile, Test1_str; " append end"
        Close #FileFile
line_17230:
Return              'february 24 2002

Do_Change_18000:
    save_line = "18000"
        'DO NOT PASS THE FIRST CHR AS A BLANK BELOW
        'THE SEARCH DOESN'T LIKE A STARTING SPACE ??
        'IN WORDPAD I WAS USING THE SEARCH AND CTRL/V
    'do not allow changes to valid e-mail files below
    If InStr(1, UCase(TheFile), "\SENT") <> 0 Then
        Print "no changes to file with \SENT*"
        GoTo line_18100
    End If
    If InStr(1, UCase(TheFile), "\INBOX") <> 0 Then
        Print "no changes to file with \INBOX*"
        GoTo line_18100
    End If
        
'    Clipboard.SetText Mid(Last_match, 2, 14)
'disabled oct 18/2000 when going back and forth it was
'nothing but a hassle maybe use elsewhere in another element.
    'put the last line of match in the clipboard
'at the office nt has "service pack 3" as well use the len below
'29 September 2003 comment out the logic below
'    If Left(os_ver, 10) = "Windows 20" And Len(os_ver) > 15 Then
'      MyAppID = Shell(Cmd(44), 3)
'    Else
      MyAppID = Shell(Cmd(10), 3)
'    End If          '30 november 2002
'29 September 2003 the above commented out

 '       AppActivate MyAppID
    SendKeys "^o", True
'    SendKeys TheFile, True 'december 3 2000
    If InStr(1, TheFile, ":") <> 0 Then
        SendKeys TheFile, True
    Else
        SendKeys App.Path + "\" + TheFile, True 'december 3 2000
    End If
    SendKeys "~", True    'FORCE AN "ENTER"
    AppActivate MyAppID 'december 27 2000
'    SendKeys (DOWN), True 'december 27 2000
'    SendKeys (Insert), True 'december 27 2000
'    SendKeys (PGDN), True 'december 27 2000
'    SendKeys "%{TAB}", True  'december 27 2000
line_18100:
'    frmproj2.Caption = program_info + " (cls #11)" '18Dec2013
    Cls         'december 27 2000
Return
Do_Append_19000:
    save_line = "19000"
    'don't allow user to append to any e-mail file below
    If InStr(1, UCase(TheFile), "\SENT") <> 0 Then
        Print "no APPENDS to file with \SENT*"
        GoTo line_19100
    End If
    If InStr(1, UCase(TheFile), "\INBOX") <> 0 Then
        Print "no APPENDS to file with \INBOX*"
        GoTo line_19100
    End If
    Open TheFile For Append As #OutFile
    If Clipboard.GetFormat(vbCFText) Then
        Clip_data = Clipboard.GetText(vbCFText)
            temps = Format(Now, "ddddd ttttt") + "       "
            temps = Left(temps, 23)     'all must be 23 long incld space
        Print #OutFile, temps, "-------------append start ---------- "
line_19005:
'
        Print #OutFile, Clip_data
        
        Print #OutFile, temps, "-------------append end   ---------- "
        Print Clip_data
        Print "Clipboard data added to "; TheFile
        Print "  === append complete ==="
        Close #OutFile
    End If
line_19100:

Return
'********************************************************
    '  * * *   E N T E R   D A T A   * * * N O T E S
'********************************************************
' enter notes data input

Do_Enter_20000:
    save_line = "20000"    'for error handling
    entered_notes = "NO"    'allow for date exitdateyes
    Enter_Count = 0

Close #OutFile
Open TheFile For Append As #OutFile
' jj logic for date here next
More_Notes_22000:
    save_line = "22000"    'for error handling
    ttt = InputBox("enter notes", "Notes Prompt   " + TheFile, , xx1 - offset1 - 2000, yy1 - offset2) '
'any 2 characters will do a paste function Ctrl/V
    'vv option will put what is in clipboard
    If InStr(1, UCase(ttt), "VVV") <> 0 Then
        ttt1 = InStr(1, UCase(ttt), "VVV")
        ttt = Left(ttt, ttt1 - 1) + ppaste + Right(ttt, Len(ttt) - (ttt1 + 2))
    End If
    If InStr(1, UCase(ttt), "GGG") <> 0 Then
        ttt1 = InStr(1, UCase(ttt), "GGG")
        ttt = Left(ttt, ttt1 - 1) + gpaste + Right(ttt, Len(ttt) - (ttt1 + 2))
    End If
'midway3 thru the program appx'==================================================================================
    
    If InStr(1, UCase(ttt), "VV") <> 0 Then
        ttt1 = InStr(1, UCase(ttt), "VV")
        ttt = Left(ttt, ttt1 - 1) + Clipboard.GetText(vbCFText) + Right(ttt, Len(ttt) - (ttt1 + 1))
    End If
    If Len(ttt) = 2 And Left(ttt, 1) = Mid(ttt, 2, 1) And UCase(ttt) <> "JJ" Then
        ttt = Clipboard.GetText(vbCFText)
  '      Print ttt 'testing only
    End If
    
    If ttt = "" And entered_notes = "YES" And _
            Cmd(7) <> "dateyes" And _
            UCase(Cmd(15)) = "EXITDATEYES" Then
            temps = Format(Now, "ddddd ttttt") + "       "
            temps = Left(temps, 23)     'all must be 23 long incld space
            Print #OutFile, temps
            Print temps
    End If
    
    If ttt = "" Then GoTo End_Notes_23000
            entered_notes = "YES"
            Enter_Count = Enter_Count + 1
            'see programmers guide page 606 for date formats below
    temps = Format(Now, "ddddd ttttt") + "       "
    temps = Left(temps, 23)     'all must be 23 long incld space
    If Cmd(7) <> "dateyes" And UCase(Cmd(14)) = "ENTERDATEYES" Then
        Cmd(14) = ""        'only do fore each new entry
        Print #OutFile, temps
        '     include "c:\dummy.txt"
   ' Insert "C:\DUMMY.BAS"
    
       Print temps
    End If
    
    If Cmd(7) <> "dateyes" Then temps = ""
    If Cmd(7) <> "dateyes" And UCase(Left(ttt, 2)) = "JJ" Then
            ttt = Right(ttt, Len(ttt) - 2)
            temps = Format(Now, "ddddd ttttt") + "       "
            temps = Left(temps, 23)     'all must be 23 long incld space
    End If
    
    
    Print #OutFile, temps & ttt
    Print temps & ttt
    GoTo More_Notes_22000
End_Notes_23000:
    If Enter_Count <> 0 Then
        Print Enter_Count; " Records added to "; TheFile
    End If
    save_line = "23000"    'for error handling
    Close #OutFile
    II = DoEvents       'yield to operating system
    
    
    GoTo What_50        'october 26 2000
'    GoTo End_32000

InputFile_24000:        'march 20/00
'    frmproj2.Caption = program_info + " (cls #12)" '18Dec2013
    Cls                 'refresh clear screen
    GoSub OpenFile_25000 'open files.txt for selection
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = True      'november 08 2001 Autoredraw pair-1
    save_line = "24000"
    For f = 1 To 20
    ttt = ""
    Line Input #FileFile, FFF
        If f = 1 Then ForeColor = QBColor(Val(Cmd(5)))
        If f = 1 Then ttt = App.EXEName + " " + vvversion + " -- " + App.Title
        If Len(FFF) > 50 Then
'derulaswamp 2010   for special autorun only         Print f, "*"; Right(FFF, 50); "   "; ttt 'december 22 2000
        If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print f, "*"; Right(FFF, 50); "   "; ttt 'december 22 2000
            Else
'derulaswamp 2010            Print f, FFF; "                            "; ttt
            If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print f, FFF; "                            "; ttt
        End If
        If f = 1 Then ForeColor = QBColor(Def_Fore)
    AllFiles(f) = FFF + ""
    Next f
    Close #FileFile
line_24005:
'*********************************************************
'main file name selection done here
'*********************************************************
'derulaswamp 2010    Print "Enter selection 1-20 e-xit or file name Option Prompt #1"
    If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print "Enter selection 1-20 e-xit or file name Option Prompt #1"
    tt1 = ""    'december 7 2000
    If screen_capture = "YES" Then
'03 September 2004        delay_sec = 5      'march 15 2001
'03 September 2004        GoSub line_30000
            new_delay_sec = 2    '03 September 2004 17Feb2017 from 5 to 2
            GoSub line_30300            '03 September 2004
    End If
    'test the getformat logic for the clipboard january 27 2002
'    If Clipboard.GetFormat(vbCFBitmap) Then
'        tt1 = InputBox("clipboard has a bitmap", , tt1, xx1 - offset1, yy1 - offset2) '
'    End If
    today_date = Format(Now, "ddddd ttttt")    'february 16 2001
'12 September 2004    If UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\" Then
'    If (UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.Path + App.EXEName), "BACKGRD") <> 0 Then
'    frmproj2.Caption = "before test " + tt1     '02 October 2004 testing
'            new_delay_sec = 2    '02 October 2004
'            GoSub line_30300            '02 October 2004
'19Aug2011    If (UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.EXEName), "BACKGRD") <> 0 Then
'12May2012 changes below
    mixx = InStr(UCase(App.EXEName), "MIX")
If mixx <> 0 Then
        tt1 = Left(App.EXEName, mixx + 2) + ".txt"
        mixx = True
        GoTo auto_p1
End If                  '30Jun2012
If Left(UCase(App.EXEName), 7) = "BIGTEXT" Then
        tt1 = "bigtext.txt"
        GoTo auto_p1
End If                  '21Jun2012
'18Jun2012  catmydrive should skip as well
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        tt1 = AllFiles(1)
        GoTo auto_p1
End If                  '18Jun2012
    
    If Left(UCase(App.EXEName), 6) = "RANDOM" Or Left(UCase(App.EXEName), 10) = "SEQUENTIAL" Then
'14May2012        tt1 = RTrim(Cmd(46))       'the command file prompt usually 1
        tt1 = AllFiles(1)   '14May2012 might want to make it file1 of the files list
'26Jun2016        tt1 = "MPEG_VIDEOS.TXT"         '24Jun2012
'06Jul2012                If mixx = True Then tt1 = "datamix.txt" '30Jun2012
    If UCase(Left(Cmd(58), 13)) = "DEFAULT_TO_CD" And Left(App.Path, 3) <> "C:\" And Len(tt1) > 2 And InStr(1, tt1, ":") <> 0 Then
        tt1 = Left(App.Path, 3) + Right(tt1, Len(tt1) - 3)
    End If          'change any C:\ to D:\ or E:\ etc
    
        GoTo auto_p1
    End If
'12May2012 changes above
    
    If (InStr(1, UCase(Cmd(45)), UCase(App.EXEName)) <> 0 And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.EXEName), "BACKGRD") <> 0 Then
        tt1 = RTrim(Cmd(46))       'the command file prompt usually 1
'-------------------------------
'02 October 2004
    If UCase(Left(Cmd(58), 13)) = "DEFAULT_TO_CD" And Left(App.Path, 3) <> "C:\" And Len(tt1) > 2 And InStr(1, tt1, ":") <> 0 Then
        tt1 = Left(App.Path, 3) + Right(tt1, Len(tt1) - 3)
    End If          'change any C:\ to D:\ or E:\ etc 02 October 2004
    
'    frmproj2.Caption = "after test " + tt1    '02 October 2004 testing
'            new_delay_sec = 2    '02 October 2004
'            GoSub line_30300            '02 October 2004
'-------------------------------
'            xtemp = InputBox(" backgrd test1", " testing Prompt #1*   ", , xx1 - offset1, yy1 - offset2)
        If debug_photo Then     '12 october 2002
            xtemp = InputBox("DOUG TESTING AUTO PROMPTS" + ttt, , , 4400, 4500) 'TESTING ONLY
        End If
        GoTo auto_p1
    End If                  '07 december 2002
    If SAVE_ttt = "in" Then tt1 = AllFiles(1)   'december 7 2000
    tt1 = InputBox("File Selection:", "File Prompt #1   " + today_date, tt1, xx1 - offset1, yy1 - offset2) '
'    Clipboard.Clear             '17Mar2016 testing only
'        Clipboard.SetText "clipboard set at prompt #1"
    If Len(tt1) > 0 And Len(tt1) < 3 Then       '19 December 2004
        xtemp = tt1
        GoSub keypad_27500
        tt1 = xtemp
    End If                                      '19 December 2004
auto_p1:
    tt1 = UCase(tt1)
    frmproj2.Caption = program_info + stretch_info       '25 november 2002
'22 december 2002
'    random_info = ""                '21 december 2002
    If Cmd(49) = "RANDOM" Then
'        random_info = " (RANDOM)"           '21 december 2002
        frmproj2.Caption = program_info + random_info + stretch_info '09 december 2002
    End If
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = False     'november 08 2001 Autoredraw pair-1
    If tt1 = "X" Or tt1 = "E" Or (tt1 = "" And SAVE_ttt <> "in") Then
        GoTo End_32000
    End If              'exit if x or e entered
    'the following 2 if statements are left in as example
    If tt1 = "25" Then
        MAX_CNT = 25
        GoTo line_24005
    End If              'internal fix till control file
    If tt1 = "20" Then
        MAX_CNT = 20
        GoTo line_24005
    End If              'internal fix till control file
'october 26 2000 the SAVE_ttt below
    If tt1 = "" And SAVE_ttt = "in" Then
        TheFile = AllFiles(1)
        tt1 = TheFile
        GoTo line_24010
    End If              'no entry default to first listed
    If tt1 = "0" Or tt1 = "1" Then
        TheFile = AllFiles(1)
        tt1 = TheFile
        SAVE_ttt = "in"     'november 03 2000 after file selected continue with defaults
        TheSearch = "."     'november 03 2000
        search_prompt = "in" 'november 03 2000
        GoTo line_24010
    End If              'no entry default to first listed
'allow for number of characters select file number KKK is 3
    If tt1 = "11" Then GoTo line_24007
    If Len(tt1) = 1 Then GoTo line_24007
    If Len(tt1) = 2 And Left(tt1, 1) <> Mid(tt1, 2, 1) Then GoTo line_24007
'need a "dbcs string manipulation function" to do following
'same as the num1$ function on the vax numeric to string
    If Left(tt1, 1) = Mid(tt1, 2, 1) Then
        tt1 = CStr(Len(tt1))
    End If

line_24007:

If Val(tt1) < 2 Or Val(tt1) > 20 Then
        GoTo line_24010
    End If
    TheFile = AllFiles(Val(tt1))
    tt1 = TheFile
    SAVE_ttt = "in"     'november 03 2000 after file selected continue with defaults
    TheSearch = "."     'november 03 2000
    search_prompt = "in" 'november 03 2000
line_24010:
    'check all file existance
    'they may have been deleted etc
    'Following line allows update if change made to control
'    If TheFile = "C:\CONTROL.TXT" Then GoSub Control_28000
    If UCase(TheFile) = "CONTROL.TXT" Then GoSub Control_28000  'december 3 2000
    save_line = "24010"
    TheFile = tt1
    LastFile = TheFile
    If save_line = "24010" Then GoTo line_24020    '29Sep2010 skip this altogether - open and read object file for defaults
    FileFile = FreeFile
    Open tt1 For Input As #FileFile
'november 3 2000 use to switch between p1 and c depending on file
    xxx_found = "NO"
line_24015:
    save_line = "24015"
    For f = 1 To 10
        Line Input #FileFile, aaa
        If Left(aaa, 4) = "xxx." Then
            xxx_found = "YES"
        End If
    Next f
    Close FileFile

line_24020:     'store new file in file.txt
    save_line = "24020"
    FileFound = 0
    For f = 1 To 20
'december 8 2000    If TheFile = AllFiles(F) Then
    If UCase(TheFile) = UCase(AllFiles(f)) Then
        FileFound = f       'save location where file is
    End If
    Next f
    If FileFound = 0 Then
        For f = 19 To 1 Step -1
        AllFiles(f + 1) = AllFiles(f)
        Next f
        AllFiles(1) = TheFile
        GoTo line_24050
    End If              'new one add it to top of list
    If FileFound > 1 Then
        For f = FileFound - 1 To 1 Step -1
        AllFiles(f + 1) = AllFiles(f)
        Next f
        AllFiles(1) = TheFile
        GoTo line_24050
    End If
line_24050:
'20 july 2002 skip the update below if cmd(40) not in file name

    If auto_exe = "YES" Then GoTo line_24080    '04 September 2004
    If InStr(Cmd(40), Left(App.Path, 3)) = 0 And Trim(Cmd(40)) <> "" Then
        If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print " no file update cmd(40)"      '14Aug2011 this was annoying
        GoTo line_24080         '20 july 2002
    End If
    save_line = "24050"
    Kill Cmd(11)        'list of files last accessed
    FileFile = FreeFile
    Open Cmd(11) For Output As FileFile
    For f = 1 To 20
        Print #FileFile, AllFiles(f)
    Next f
line_24080:
    Close FileFile
    save_line = "24080"         '12Jun2016
line_24090: 'return exit line
    If InStr(UCase(TheFile), ".MBX") <> 0 And SAVE_ttt = "in" Then mbxyes = "Y" 'december 17 2000

Return

OpenFile_25000:         'march 20/00
'files.txt is the last 20 opened files
'for display and selection
    save_line = "25000"
        FileFile = FreeFile
        Open Cmd(11) For Input As #FileFile
    
Return
hilite_25500:
'        xtemp = InputBox(CStr(hilite_cnt) + yyy + hilite_hh + ooo, "Continue Prompt", , 2000, 2000)

    If InStr(1, UCase(ooo), "XXX.") <> 0 Then
    hilite_cnt = hilite_cnt + 1
    cript2(hilite_cnt) = Mid(ooo, InStr(1, UCase(ooo), UCase(hilite_this)) + Len(hilite_this))
'        tt1 = InputBox("testing prompt" + cript2(hilite_cnt), , , 4400, 4500) 'TESTING ONLY
    End If
Return              'april 22 2001

Search_26000:        'march 20/00
'        xtemp = InputBox("(A) testing prompt 02 January 2005" + sscreen_saver_ww + inin, , , 4400, 4500) '02 January 2005
    If tempdata = "COMMAND" Then GoTo line_26002a           '02 January 2005
    If sscreen_saver_ww = "YES" And inin <> "" Then GoTo line_26090 '28 april 2002
'    frmproj2.Caption = program_info + " (cls #13)" '18Dec2013
    Cls                 'refresh clear screen
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = True      'november 08 2001 autoredraw pair-3

            If Not mixx Then Set Picture = LoadPicture()           '12Feb2017
'    Set Picture = LoadPicture() 'clear any picture lp#7   normal   12Feb2017
    If debug_photo Then         '12 october 2002
        tt1 = InputBox("testing photo 3.3", , , 4400, 4500)  'TESTING ONLY
    End If
    GoSub OpenFile_27000 'open search.txt for selection
    save_line = "26000"
    For f = 1 To 20
    
    Line Input #FileFile, FFF
    If TheSearch = "." Then
'december 7 2000        If F = 1 Then ForeColor = QBColor(Val(Cmd(5)))
'derulaswamp 2010        Print f, FFF        'DISP OPTIONS IF "." ENTERED
        If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print f, FFF        'DISP OPTIONS IF "." ENTERED
'december 7 2000        If F = 1 Then ForeColor = QBColor(Def_Fore)
    End If
    AllSearch(f) = FFF + ""
    Next f
    Close #FileFile
    If TheSearch <> "." Then
        GoTo line_26020
    End If
'*************************************************
' the search selection 1-20 made here
'*************************************************
    If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print "Enter selection 1-20 e-xit a for all or new search"
line_26002:
    If ttt = "." Then   'january 05 2001
        ttt = AllSearch(1)
    Else
        ttt = ""            'december 7 2000
    End If          'january 05 2001
    If search_prompt = "in" Then ttt = Cmd(33)  'december 7 2000
    If emailsea = "Y" And search_prompt = "in" Then ttt = "a"        'december 11 2000
    Cmd(33) = ""    'december 7 2000
    If screen_capture = "YES" Then
'03 September 2004        delay_sec = 5      'march 15 2001
'03 September 2004        GoSub line_30000
            new_delay_sec = 5    '03 September 2004
            GoSub line_30300            '03 September 2004
    End If
'option prompt #3
'12 September 2004    If UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\" Then
'    If (UCase(App.EXEName) = Trim(UCase(Cmd(45))) And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.Path + App.EXEName), "BACKGRD") <> 0 Then
'19Aug2011
    '12May2012
    If Left(UCase(App.EXEName), 6) = "RANDOM" Or Left(UCase(App.EXEName), 10) = "SEQUENTIAL" Then
'14May2012        ttt = "PHOTO"
'24Jun2012        ttt = AllSearch(1) '14May2012 change this to the last search string
        ttt = "PHOTO"       '24Jun2012
'26Jun2016        GoTo auto_p3                '12May2012
    End If
    If (InStr(1, UCase(Cmd(45)), UCase(App.EXEName)) <> 0 And Left(App.Path, 3) <> "C:\") Or InStr(1, UCase(App.EXEName), "BACKGRD") <> 0 Then
'            xtemp = InputBox(" backgrd test3", " testing Prompt #3*   ", , xx1 - offset1, yy1 - offset2)
        ttt = RTrim(Cmd(48))       'the search string prompt usually photo
        GoTo auto_p3        'make sure that cd search file does not have "C:\"
                            'it should be "D:\" anything but "C:\" to work
    End If                  '07 december 2002
'19 September 2004  command== stuff
    If Len(command_line) > 1 Then
        'i do not think the file is open here it should be 1 line past the match for now...
        'maybe skip the multiple reads below if random is on...????
'        temptemp = InputBox("DOUG TESTING command_linea==" + CStr(zzz_cnt) + " " + aaa, , , 4400, 4500) 'TESTING ONLY
        command_line = ""
'        temptemp = InputBox("DOUG TESTING command_lineb==" + command_line, , , 4400, 4500) 'TESTING ONLY
        multi_prompt2 = ""  '20 September 2004
'        aaa = ""            '20 September 2004
'        GoTo input_1000a     '20 September 2004
        For zzz_cnt = 1 To hold_zzz
        Line Input #OutFile, aaa '21 September 2004 27Aug2010 might need some changes here?
        Next zzz_cnt                '21 September 2004
            ooo = aaa           '21 September 2004
'        temptemp = InputBox("DOUG TESTING command_lineaa==" + CStr(hold_zzz) + " " + aaa, , , 4400, 4500) 'TESTING ONLY
        GoTo input_1000b     '21 September 2004
'        ttt = RTrim(Cmd(48))       '20 September 2004  this changed things a bit???
'        GoTo Line_17055
'        GoTo auto_p3         '21 September 2004
    End If                  '19 September 2004
        If debug_photo Then     '12 october 2002
'            xtemp = InputBox("DOUG TESTING auto prompt #3" + ttt, , , 4400, 4500) 'TESTING ONLY
        End If
'26 November 2004 skip around this prompt is a control file switch took place last
'            xtemp = InputBox("DOUG TESTING auto prompt #3 search=" + temptemp + " tempdata=" + tempdata + " " + eofsw, , , 4400, 4500) 'TESTING ONLY
    If eofsw = "YES" Then
        eofsw = ""              '26 November 2004
        ttt = Trim(temptemp)
        GoTo auto_p3
    End If                      '26 November 2004
'            xtemp = InputBox("testing 02 January 2005=" + temptemp + " tempdata=" + tempdata + " " + eofsw, , , 4400, 4500) 'TESTING ONLY
line_26002a:                        '02 January 2005
    If tempdata = "CONTROL" Or tempdata = "COMMAND" Then         '26 November 2004
'            xtemp = InputBox("(B) testing 02 January 2005=" + temptemp + " tempdata=" + tempdata + " " + eofsw, , , 4400, 4500) 'TESTING ONLY
        ttt = Trim(temptemp)               '26 November 2004 this will change to something photo or "a" etc
        tempdata = ""
        GoTo auto_p3
    End If                          '26 November 2004
'21Jun2012 check
If mixx = True Then
        ttt = "PHOTO"
        GoTo auto_p3
End If                  '30Jun2012b
If Left(UCase(App.EXEName), 7) = "BIGTEXT" Then
        ttt = "A"
        GoTo auto_p3
End If                  '21Jun2012
    ttt = InputBox("<" + UCase(prompt2) + "> Search Selection:", "Search Prompt #3    " + TheFile, ttt, xx1 - offset1, yy1 - offset2) '
'19 December 2004
    If Len(tt1) > 0 And Len(tt1) < 3 Then       '19 December 2004
        xtemp = ttt
        GoSub keypad_27500
        ttt = xtemp
    End If                                      '19 December 2004
    If show_files_yn Then frmproj2.Caption = program_info + random_info + stretch_info '24 december 2002
auto_p3:
'    frmproj2.Caption = program_info + " (cls #14)" '18Dec2013
    Cls                             'november 10 2001 needed with autoredraw
'        temptemp = InputBox("DOUG TESTING command_lineb==" + ttt, , , 4400, 4500) 'TESTING ONLY
    If auto_redraw = "YES" Then frmproj2.AutoRedraw = False      'november 08 2001 autoredraw pair-3
'        xtemp = InputBox("TESTING DOUG" + ttt + "*" + search_prompt, , , 4400, 4500) 'TESTING ONLY
'november 6 2000
'    Me.MousePointer = vbHourglass       '18 august 2002
parsez:      'january 12 2001
    If Left(ttt, 2) = "  " Then
        ttt = Mid(ttt, 2)
        GoTo parsez
    End If
    i = InStr(ttt, "  ")
    If i = 0 Then GoTo noparsez
    ttt = Left(ttt, i - 1) + Mid(ttt, i + 1)
    GoTo parsez
noparsez:
    
    Context_cnt = -1     'november 10 2000
'17 December 2004   add in the V or H prompt for insert
    If UCase(ttt) = "V" Or UCase(ttt) = "H" Then
         If Clipboard.GetFormat(vbCFText) Then
             ttt = Clipboard.GetText(vbCFText)
         End If          '17 December 2004
    End If              '17 December 2004
    If UCase(ttt) = "F" Then
        prompt2 = "C"
        SAVE_ttt = "F"
        mbxi = 1        'january 31 2001
        GoTo line_26002
    End If              'january 31 2001
    If UCase(ttt) = "Q" Then
        prompt2 = "Q"
        SAVE_ttt = "Q"
        mbxi = 1        'december 17 2000
        GoTo line_26002
    End If              'december 2 2000
    
    
    If UCase(ttt) = "C" Then
        prompt2 = "C"
        SAVE_ttt = "C"
        mbxi = 1        'december 29 2000
        GoTo line_26002
    End If              'december 2 2000
    If prompt2 = "Q" And ttt <> "" Then
        qqq = ttt + ""
    End If
    If prompt2 = "Q" And ttt = "" Then
        qqq = ""
    End If      'november 6 2000
'november 6 2000
'    If prompt2 <> "Q" And ttt <> "d" Then
'        ttt = UCase(ttt)    'november 25 2000
'    End If
'august 27/00
    If ttt = "E" Or ttt = "e" Or ttt = "X" Or ttt = "x" Or ttt = "" Then
        GoTo line_26010
    End If
    If ttt = "" Or ttt = "0" Or ttt = "1" Then
        TheSearch = AllSearch(1)
        ttt = TheSearch
        qqq = TheSearch
        GoTo line_26010
    End If              'no entry default to first listed
'november 14 2000 the following stuff moved up to do_search area
'    If UCase(ttt) = "XXX" Then
'        extract_yes = "YES"     'november 12 2000
'            ExtFile = FreeFile
''            Kill Cmd(19)
'            DoEvents
'     Open Cmd(19) For Output Access Write As #ExtFile
''           Open Cmd(19) For Output As #ExtFile
''            Print #ExtFile, "testing output"
'            DoEvents
'            line_len = 500 'on extract do not do any wraps???
'            GoTo line_26002
'    End If
    
'allow string length to determine search selection
'december 28 2000 comment out the 6 following lines
'    If ttt = "11" Then GoTo line_26007
'    If Len(ttt) = 1 Then GoTo line_26007
'    If Len(ttt) = 2 And Left(ttt, 1) <> Mid(ttt, 2, 1) Then GoTo line_26007
'    If Len(ttt) >= 2 And Left(ttt, 1) = Mid(ttt, 2, 1) And Mid(ttt, 2, 1) = Mid(ttt, 3, 1) Then
'        ttt = CStr(Len(ttt))
'    End If
line_26007:
    save_line = "26007"     'december 28 2000
    If Len(ttt) = 2 And ttt >= "10" And ttt < "21" Then GoTo line_26008
    If Len(ttt) = 1 And ttt > "0" And ttt <= "9" Then GoTo line_26008
'    ttt = ""        'january 05 2001
    GoTo line_26010     'december 28 2000
    'did the code above so no numeric error trap would be needed

line_26008:
'december 28 2000 commented out the 3 following lines
'    If Val(ttt) < 2 Or Val(ttt) > 20 Then
'        GoTo line_26010
'    End If
    TheSearch = AllSearch(Val(ttt))
    ttt = TheSearch
    qqq = TheSearch
line_26010:
    TheSearch = ttt
'31 august 2002 ucase stuff below
'when all are numeric ie shifting cases makes no difference don't use uppercase otherwise do ie "Y"
    do_tab = 0    '05 october 2002 no need to check for tabs if search has no spaces
    If InStr(ttt, " ") > 0 Then do_tab = True '05 october 2002
    If UCase(ttt) = ttt And LCase(ttt) = ttt Then
        uppercase = "N"
'        temptemp = InputBox("TESTING ucase no" + prompt2 + p2p2 + SAVE_ttt + ttt + "*" + UCase(ttt) + "*" + LCase(ttt), , , 4400, 4500) 'TESTING ONLY
    End If
    If UCase(ttt) <> ttt Or LCase(ttt) <> ttt Then
        uppercase = "Y"
'        temptemp = InputBox("TESTING ucase yes" + prompt2 + p2p2 + SAVE_ttt + ttt + "*" + UCase(ttt) + "*" + LCase(ttt), , , 4400, 4500) 'TESTING ONLY
    End If
'31 august 2002
line_26020:     'store new search in search.txt
    save_line = "26020"
    FileFound = 0
    For f = 1 To 20
    If TheSearch = AllSearch(f) And Len(TheSearch) > 0 Then
        FileFound = f       'save location where search is
    End If
    Next f
'        Print "TheSearch="; TheSearch
'        tt1 = InputBox("continue", , , 4400, 4500)  'TESTING ONLY

    If UCase(TheSearch) = "D" Then
        TheSearch = ""  'do not log D day searches
    End If
    If UCase(TheSearch) = "M" Then
        TheSearch = ""  'do not log M month searches
    End If
    If UCase(TheSearch) = "DD" Then
        TheSearch = ""  'do not log DD day searches
    End If
    If UCase(TheSearch) = "MM" Then
        TheSearch = ""  'do not log MM month searches
    End If
    If FileFound = 0 And TheSearch <> "" Then
        For f = 19 To 1 Step -1
        AllSearch(f + 1) = AllSearch(f)
        Next f
        AllSearch(1) = TheSearch
        GoTo line_26050
    End If              'new one add it to top of list
    If FileFound > 1 Then
        For f = FileFound - 1 To 1 Step -1
        AllSearch(f + 1) = AllSearch(f)
        Next f
        AllSearch(1) = TheSearch
        GoTo line_26050
    End If
line_26050:
    save_line = "26050:"
    If auto_exe = "YES" Then GoTo line_26080    '04 September 2004
    If InStr(Cmd(40), Left(App.Path, 3)) = 0 And Trim(Cmd(40)) <> "" Then
'07Aug2012 no need for the following line message
'07Aug2012        If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print " no search file update cmd(40)"      '14Aug2011 this was annoying
'            xtemp = InputBox("app.path" + App.Path + " not found or updated " + Cmd(40), , , 4400, 4500) 'TESTING ONLY
        GoTo line_26080
    End If              '20 july 2002 must be a writeable device in cmd(40)
    Kill Cmd(12)        'search strings save file
    FileFile = FreeFile
    Open Cmd(12) For Output As FileFile
    For f = 1 To 20
        Print #FileFile, AllSearch(f)
    Next f
line_26080:     '20 july 2002
    Close FileFile
line_26090: 'return exit line
    TheSearch = ""
Return

OpenFile_27000:         'march 20/00
'search.txt is the last 20 search strings
'for display and selection
    save_line = "27000"
        FileFile = FreeFile
        Open Cmd(12) For Input As #FileFile
    
Return

keypad_27500:           '19 December 2004
    xtemp = UCase(xtemp)
    If xtemp = "M" Then xtemp = "1"
    If xtemp = "," Then xtemp = "2"
    If xtemp = "." Then xtemp = "3"
    If xtemp = "J" Then xtemp = "4"
    If xtemp = "K" Then xtemp = "5"
    If xtemp = "L" Then xtemp = "6"
    If xtemp = "U" Then xtemp = "7"
    If xtemp = "I" Then xtemp = "8"
    If xtemp = "O" Then xtemp = "9"
    If xtemp = "M " Then xtemp = "10"
    If xtemp = "MM" Then xtemp = "11"
    If xtemp = "M," Then xtemp = "12"
    If xtemp = "M." Then xtemp = "13"
    If xtemp = "MJ" Then xtemp = "14"
    If xtemp = "MK" Then xtemp = "15"
    If xtemp = "ML" Then xtemp = "16"
    If xtemp = "MU" Then xtemp = "17"
    If xtemp = "MI" Then xtemp = "18"
    If xtemp = "MO" Then xtemp = "19"
    If xtemp = ", " Then xtemp = "20"
Return                  '19 December 2004
'control.txt control.txt control.txt information read into cmd elements here read control file ***vip***
Control_28000:      'april 20/00 update control file info read it in
    save_line = "28000"
    If InStr(UCase(aaa), "CONTROL==") <> 0 Then GoTo line_28005   '19 November 2004
    If InStr(Cmd(40), Left(App.Path, 3)) = 0 And Trim(Cmd(40)) <> "" And SAVE_ttt <> "in" Then
'            xtemp = InputBox("app.path" + App.Path + " control file not re-read " + Format(delay_sec, "###0.0000") + Cmd(40) + SAVE_ttt, , , 4400, 4500) 'TESTING ONLY
'        Print " no control file reread cmd(40)"
           frmproj2.Caption = " no control file reread cmd(40)" '19 November 2004 testing only
        GoTo line_28099
    End If              '20 july 2002 must be a writeable device in cmd(40)
line_28005:         '19 November 2004 allow re-read if in control file as it is just switching files
'            xtemp = InputBox("app.path" + App.Path + " control file not re-read " + Format(delay_sec, "###0.0000") + Cmd(40) + SAVE_ttt, , , 4400, 4500) 'TESTING ONLY
'    FileFile = FreeFile
    CtrlFile = FreeFile
'    Open "C:\control.txt" For Input As #FileFile
    Open control_file For Input As #CtrlFile
line_28010:
    save_line = "28010"
    For f = 1 To 100
    Line Input #CtrlFile, FFF
    Cmd(f) = FFF
    Next f
    
'all new elements should be read in and set to values here
'eventually moving them from the prompt area to here.

'08 January 2004 slomo_seg
    slomo_seg = 0
    slomo_seg = Val(Cmd(62))
    If slomo_seg = 0 Then
'21 March 2004        slomo_seg = 0.01667
        slomo_seg = 0.00668
    End If
'08 January 2004
    
    pad_time = 0                '21 March 2004
    pad_time = Val(Cmd(64))      '21 March 2004
    If pad_time = 0 Then
'12Jan2018 disable        pad_time = 30
    End If                      '21 March 2004
    
'04 November 2003 have play_speed in control file
    play_speed = 1000
    If Left(UCase(Cmd(61)), 8) = "SETSPEED" Then
        f = Len(Cmd(61)) - 8
        play_speed = Val(Right(Cmd(61), f))
        If play_speed = 0 Then
            play_speed = 1000
        End If
    End If              '04 November 2003
    save_play_speed = play_speed        '16 November 2003
'18 March 2003 ver=1.01
    If Left(Cmd(55), 2) > "00" And Left(Cmd(55), 2) <= "99" Then
        Clear_Context_lines = Val(Left(Cmd(55), 2))
    Else
        Clear_Context_lines = 0     '18 March 2003 ver=1.01
    End If
'18 March 2003 ver=1.01
        
'22 March 2003 ver=1.02
    If UCase(Cmd(56)) = "PHOTO_DETAIL" Then
        Cmd(56) = "PHOTO_DETAIL"        'just to make sure all capitals
        detailyn = "PHOTO_DETAIL"       '18 November 2004
    Else
        detailyn = "noPHOTO_DETAIL"     '18 November 2004
    End If
'22 March 2003 ver=1.02

    If UCase(Cmd(53)) = "SHOWVIDEO" Then
        videoyn = "SHOWVIDEO"
    Else
        videoyn = "NOSHOWVIDEO"
    End If                      '10 FEBRUARY 2003
    If Left(UCase(Cmd(50)), 9) = "SHOWFILES" Then
        show_files_yn = True '24 december 2002
'       xtemp = InputBox("DOUG 3 show_files_yn " + Cmd(50), , , 4400, 4500) 'TESTING ONLY
    End If
'19Aug2011
    If Left(UCase(App.EXEName), 6) = "RANDOM" Or Left(UCase(App.EXEName), 10) = "SEQUENTIAL" Then
        Cmd(49) = "NORANDOM"
        If Left(UCase(App.EXEName), 6) = "RANDOM" Then
            Cmd(49) = "RANDOM"
'       xtemp = InputBox("testing random set " + Cmd(50), , , 4400, 4500) 'TESTING ONLY
        End If
    End If      '12May2012
   
    
    If (InStr(1, UCase(Cmd(45)), UCase(App.EXEName))) <> 0 Then auto_exe = "YES"  '07 december 2002
'==============================================================
    random_info = ""            '21 december 2002
    If Left(UCase(Cmd(49)), 6) = "RANDOM" Then
        random_info = " (RANDOM)"   '21 december 2002
        rand = -1           'when multiples allowed as below remove this line

'19 august 2003 try skipping the following line... for now
'19 august 2003        If xxx_found = "NO" And text_pause <> True Then rand = 0   '19 january 2003
'       xtemp = InputBox("testing douga 19 august 2003 " + Cmd(50), , , 4400, 4500) 'TESTING ONLY
'19 august 2003        If xxx_found = "NO" Then rand = 0   '19 january 2003
    End If              '09 december 2002
'19 August 2003    If Cmd(49) <> "RANDOM" Then rand = 0  '09 december 2002
    If Left(Cmd(49), 6) <> "RANDOM" Then rand = 0  '09 december 2002
'=============================================================
    rand1 = 0           '23 March 2004
    If Left(UCase(Cmd(65)), 7) = "RANDBEG" Then
        rand1 = -1
    End If              '23 March 2004
    
    thumb_nail = "NO"       '14 April 2004
    If Left(UCase(Cmd(67)), 5) = "THUMB" Then
        thumb_nail = "YES"
    End If                  '14 April 2004
    
    save_line = "28020"
        Def_Fore = 15       'assign foreground to white
    Def_Fore = Cmd(4)
    Hold_Fore = Def_Fore    '27 July 2003
    sep = Cmd(6)        'search string seperator eg "/" or "." etc
    Set_Fore = 12      'assign set to red
    Set_Fore = 14        'try yellow
    Set_Fore = Cmd(5)
    line_len = Val(Cmd(21))
    over_lap = Val(Cmd(13))     'january 10 2001
'      xtemp = InputBox("testing prompt overlap=" + CStr(over_lap) + " " + CStr(wrap_cnt) + " " + CStr(cnt), "test", , 4400, 4500)  '
    If over_lap < 1 Then over_lap = 10   'january 10 2001
    If Val(line_len) < 10 Then
        line_len = 82
    End If
    Context_lines = Val(Cmd(22))
    If Context_lines > 40 Then Context_lines = 40   'February 04 2001
    If Context_lines < 1 Then
        Context_lines = 10
    End If
    photo_copy = Cmd(23)    'march 17 2001
    If Len(photo_copy) < 7 Then
        photo_copy = "d:\search\tempfold\"
    End If      'march 17 2001
    ppaste = Cmd(24)    'allow for vvv paste in data entry
    gpaste = Cmd(25)    'allow for ggg paste in data entry
    ss_search = Cmd(26)
    If ss_search = "" Then
        ss_search = "PHOTO"
        screensave(1) = "PHOTO"
    End If
    delay_sec = Val(Cmd(27)) 'wait pause timing
    smlbud = Val(Cmd(85)) 'budge amount for video offset 17Dec2017
    bigbud = Val(Cmd(86)) 'the skip amount in fast forward and budge in effect 17Dec2017
    maxbud = Val(Cmd(87)) ' the total number of budge allowed 17Dec2017
'    If smlbud = 0 Then smlbud = 2  'budge amount not set in older control files 17Dec2017
    hold_sec = delay_sec    '22 March 2004
'demo don't allow them the privilege of changing display time
    If ddemo = "YES" Then
       delay_sec = 0
    End If
    If delay_sec < 0 Then
        delay_sec = 4       '08 July 2003 what the heck is this for ***vip*** check it out
    End If
        xx1 = 8000
        yy1 = 6500
'        If Test1_str = "P1" Then
            xx1 = Val(Cmd(17))
            yy1 = Val(Cmd(18))
'        End If
        If xx1 < 100 Then
            xx1 = 8000
        End If
        If yy1 < 100 Then
            yy1 = 6500
        End If
    
    GoSub line_29100    'noshow routine
    GoSub line_29200    'screen saver routine
    ForeColor = QBColor(Def_Fore)
    BackColor = QBColor(3)
    BackColor = QBColor(Val(Cmd(3)))
'    BackColor = QBColor(Rnd * 20) 'add some color
                'it randomizes to black and can't see a thing??
    MAX_CNT = 20    'IF FONT.SIZE CHANGES SO DOES THIS COUNT
    MAX_CNT = Val(Cmd(1))
    Font.Size = 12
    Font.Size = Val(Cmd(2))
    photo_copy = Cmd(23) '14Jul2016
'    photo_copy = Val(Cmd(23)) 'march 17 2001
'    If back_cnt < 1 Then
'        back_cnt = MAX_CNT
'    End If
    If Cmd(29) = "" Then
        Cmd(29) = "12"
    End If
    If Val(Cmd(29)) < 1 Then
        Cmd(29) = "12"
    End If              'October 27 2000
    AltColor = Val(Cmd(29))
    If Cmd(30) <> "N" Then
        Cmd(30) = "Y"
    End If              'October 27 2000
    
    'this was an attempt to get the problem with the Millennium ME software
 '   If UCase(Cmd(30)) <> "Y" Then Cls 'december 8 2000
 '   If UCase(Cmd(30)) <> "Y" Then frmProj2.BorderStyle = 2 'december 8 2000
 '   If UCase(Cmd(30)) <> "Y" Then frmProj2.MaxButton = True 'december 8 2000
 '   If UCase(Cmd(30)) <> "Y" Then frmProj2.Caption = "Millennium" 'december 8 2000
 '   If UCase(Cmd(30)) <> "Y" Then frmProj2.Height = 9005 'december 8 2000
    hilite_this = Cmd(31)   ' only hilites data not on matching line
    If Len(Cmd(32)) < 7 Then
'        Cmd(32) = "c:\cript.txt"
        Cmd(32) = "cript.txt"   'december 3 2000
    End If              'november 20 2000
    cript_file = Cmd(32)
    If Cmd(33) = "     " Then Cmd(33) = ""  'december 7 2000
    context_win = Val(Cmd(34))              'january 01 2001
    If UCase(Cmd(38)) = "STRETCH" Then
        stretch_info = " (STRETCH)"         '21 december 2002
        stretch_img = "YES"
        img_ctrl = "YES"
    Else
        stretch_info = " (NORMAL)"              '21 december 2002
        stretch_img = "NO"
        img_ctrl = "NO"
    End If                      'september 23 2001
    auto_redraw = "NO"         'november 10 2001
    frmproj2.AutoRedraw = False      'november 10 2001 autoredraw
    If UCase(Cmd(39)) = "AUTOREDRAW" Then
        auto_redraw = "YES"
        frmproj2.AutoRedraw = True      'november 10 2001 autoredraw
    End If                      'november 10 2001
        adjust_sec = Val(Cmd(59))    '29 September 2003
        freeze_sec = Val(Cmd(60))       '01 October 2003
        OffSet = 0                      '02Nov2011
    If Left(UCase(Cmd(83)), 8) = "OFFSET==" Then
        f = Len(Cmd(83)) - 8
        OffSet = Val(Right(Cmd(83), f))
    End If              '02Nov2011
 
'        xtemp = InputBox("DOUG TESTING " + auto_redraw + Cmd(39) + "*", , , 4400, 4500) 'TESTING ONLY
    If context_win < 3 Then context_win = 3   'january 01 2001
    If context_win > MAX_CNT Then context_win = MAX_CNT 'january 01 2001
line_28099:         '20 july 2002
    Close CtrlFile
Return

replace_29000:  'August 10/00 search / replace sequence
save_line = "29000"
'    frmproj2.Caption = program_info + " (cls #15)" '18Dec2013
    Cls     'clear screen
    If encript <> "RRR" Then
        GoTo line_29000a
    End If      'november 20 2000
    criptcnt = 0
line_29000a:
    xtemp = Cmd(19)   'usually replace.txt
    GoSub line_16000    'get file name
'            xtemp = InputBox("27 october 2004 TESTING 5b =" + encript + "=", , , 4400, 4500) 'TESTING ONLY
    temp2 = 0
'    DoEvents        '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 5b1 =" + encript + "=", , , 4400, 4500) 'TESTING ONLY
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
        OutFile = FreeFile              '27 October 2004 this may be a problem for just the RRR function
    End If                              '27 October 2004   put the if around it just so things do not change
    Open TheFile For Input As #OutFile
'           xtemp = InputBox("27 october 2004 TESTING 5b2 =" + encript + "=", , , 4400, 4500) 'TESTING ONLY
    If encript <> "RRR" Then
        GoTo line_29005
    End If      'november 20 2000
     If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
        Line Input #BatchFile, case_yes
'            xtemp = InputBox("27 october 2004 TESTING 5c =" + case_yes + "=", , , 4400, 4500) 'TESTING ONLY
        GoTo skip_case
    End If                                              '27 October 2004
    case_yes = UCase(InputBox("Case sensitive change Y/N <N>", "Case sensitive Prompt", , xx1 - offset1, yy1 - offset2))
skip_case:                                              '27 October 2004
    If case_yes <> "Y" Then
        case_yes = "N"
    End If
line_29003:
    If case_yes = "Y" Then
        If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 6 " + ttt, , , 4400, 4500) 'TESTING ONLY
            Line Input #BatchFile, in_str
            GoTo skip_in_str
        End If
            in_str = InputBox("String to change", "From string Prompt", , xx1 - offset1, yy1 - offset2)
skip_in_str:                        '27 October 2004
    End If
    If case_yes <> "Y" Then
        If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 7 " + ttt, , , 4400, 4500) 'TESTING ONLY
            Line Input #BatchFile, in_str
'            in_str = UCase(in_str)
            GoTo skip_in_str1
        End If
            in_str = UCase(InputBox("String to change", "From string Prompt", , xx1 - offset1, yy1 - offset2))
skip_in_str1:                       '27 October 2004
    End If
'30 October 2004 this worked for blank lines comming in in the batch file. Now it does not do much else on blanks
'30 October 2004    If in_str = "" Or Left(UCase(in_str), 4) = "END*" Then
'03 November 2004    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" And in_str = "" Then in_str = " "   '30 October 2004
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then
        temptemp = in_str           '03 November 2004
check_tab:                          '03 November 2004
        tt = InStr(temptemp, Chr(9)) 'check for tabs
        If tt = 0 Then
            GoTo past_tab
        End If
        in_str = Left(temptemp, tt - 1) + "    " + Mid(temptemp, tt + 1) 'replace tab with spaces (consistancy of code too)
        GoTo check_tab
past_tab:                           '03 November 2004
        'temptemp needed here so that tabs elsewhere are not removed (only if tabs and spaces only on line)
        If Trim(in_str) = "" And in_str = temptemp Then
'           frmproj2.Caption = " trim found " + CStr(criptcnt) '03 November 2004 testing only
'            new_delay_sec = 5      '03 November 2004  testing only
'            GoSub line_30300        '03 November 2004  testing only
        GoTo line_29003   '03 November 2004
        End If          '03 November 2004
    End If
    If in_str = "" Or Left(UCase(in_str), 4) = "END*" Then
'            xtemp = InputBox("27 october 2004 TESTING 7a " + in_str, , , 4400, 4500) 'TESTING ONLY
        GoTo line_29005
    End If
    Print in_str; " ";
'    ttt1 = Len(in_str)
        If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then    '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 8 " + ttt, , , 4400, 4500) 'TESTING ONLY
'            Line Input #BatchFile, out_str
'03 November 2004            out_str = "nothing"             '27 October 2004 testing only
            out_str = " "    '03 November 2004
            GoTo skip_out_str
        End If                                      '27 October 2004
    out_str = InputBox("New string", "To string Prompt", , xx1 - offset1, yy1 - offset2)
skip_out_str:                       '27 October 2004
'            xtemp = InputBox("27 october 2004 TESTING 8a *" + out_str + "*", , , 4400, 4500) 'TESTING ONLY
    If out_str = "" Or UCase(out_str) = "END*" Then
'            xtemp = InputBox("27 october 2004 TESTING 8b " + ttt, , , 4400, 4500) 'TESTING ONLY
        GoTo line_29005
    End If
    Print out_str
    If criptcnt Mod 20 = 0 Then
            frmproj2.Caption = " working- " + CStr(criptcnt) '28 October 2004
        Cls
    End If          '27 October 2004 clear the screen only
    criptcnt = criptcnt + 1
    cript1(criptcnt) = in_str
    cript2(criptcnt) = out_str
    cript3(criptcnt) = Len(cript1(criptcnt))
'            xtemp = InputBox("27 october 2004 TESTING 7aa " + CStr(criptcnt), , , 4400, 4500) 'TESTING ONLY
'    II = Len(out_str)
'    If Len(Cmd(19)) < 7 Then
'        Cmd(19) = "c:\replace.txt"
'    End If
    GoTo line_29003

line_29005:
'                xtemp = InputBox("27 october 2004 TESTING 9 " + ttt, , , 4400, 4500) 'TESTING ONLY
    If criptcnt = 0 Then
        GoTo line_29095
    End If

    save_line = "29005"
    Kill FileExt
line_29008:
    FileFile = FreeFile
    Open FileExt For Output Access Write As #FileFile
'    criptcnt = 1        'november 18 2000
'    cript1(1) = in_str  'november 18 2000
'    cript2(1) = out_str 'november 18 2000
    dblStart = Timer      'get the start time
    zzz_cnt = 0
line_29010:
    save_line = "29010"
    Line Input #OutFile, aaa
line_29015:         'allow for the replacement right here february 10 2001
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then GoTo line_29020   '27 October 2004
    tt = InStr(aaa, Chr(9)) 'check for tabs
    If tt = 0 Then
        GoTo line_29020
    End If
    'change any tabs to 4 spaces
    aaa = Left(aaa, tt - 1) + "    " + Mid(aaa, tt + 1)
    GoTo line_29015             'february 10 2001 if search and replace done do tabs
line_29020:                     'february 10 2001
    save_line = "29020"         '17Dec2017
    If encript = "RRR" Then
        GoSub line_29300       'do line at a time november 18 2000
    Else
    zzz_cnt = zzz_cnt + 1
'    If zzz_cnt Mod 10000 = 0 Then
    If zzz_cnt Mod 10000 = 0 Then
        DoEvents        'december 06 2001
        Print "working "; zzz_cnt; Format(Now, "ddddd ttttt")
    End If
        GoSub line_29400       'do line at a time november 18 2000
    End If
        yyy_cnt = yyy_cnt + 1       '28 October 2004
        If yyy_cnt Mod 1000 = 0 Then        '17dec2017
'        If yyy_cnt Mod 100 = 0 Then
                frmproj2.Caption = " working at " + CStr(yyy_cnt) + " of " + CStr(criptcnt) '28 October 2004
        End If          '27 October 2004
    GoTo line_29010
line_29090:
    dblEnd = Timer      'get the end time
    Print "     elap="; Format(dblEnd - dblStart, "#####0.000")
    Print "rename "; FileExt; " as "; TheFile
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then        '27 October 2004
            frmproj2.Caption = " total matches " + CStr(criptcnt) '27 October 2004
            new_delay_sec = 3      '27 October 2004  testing only
            GoSub line_30300        '27 October 2004  testing only
        tt1 = "N"
        GoTo skip_rename
    End If                                                  '27 October 2004
    tt1 = UCase(InputBox("total changes=" + CStr(temp2) + " rename files Y/N <Y>", "Rename Prompt", , xx1 - offset1, yy1 - offset2)) 'TESTING ONLY
skip_rename:                                                '27 October 2004
    If tt1 <> "N" Then
        tt1 = "Y"
    End If
    If UCase(tt1) = "N" Then
        Close FileFile, OutFile
        GoTo line_29095
    End If
    Close FileFile, OutFile
    DoEvents
    save_line = "29092"
'    Kill "c:\oldfile.txt"
    Kill "oldfile.txt"  'december 3 2000
    DoEvents
line_29093:
'     tt1 = InputBox("minor pause", , , 4400, 4500) 'TESTING ONLY
    DoEvents
'    Name TheFile As "c:\oldfile.txt"
    Name TheFile As "oldfile.txt"   'december 3 2000
    DoEvents
    Name FileExt As TheFile
    'rename files here if Y entered
line_29095:
    DoEvents
Return

line_29100:     'noshow elements set up here
    save_line = "29100"
    For II = 1 To 10
        noshow(II) = ""
    Next II
    If Len(Cmd(28)) < 3 Then
        GoTo line_29190
    End If
    nocount = 0
    temps = Cmd(28) + ""
line_29150:
    save_line = "29150"
    tt = InStr(temps, " ")
    If tt = 1 Then
        temps = Right(temps, Len(temps) - 1)
        GoTo line_29150 'strip out any leading spaces
    End If
    If tt = 0 Then
        If Len(temps) > 2 Then
            nocount = nocount + 1
            noshow(nocount) = temps
 '       Print "noshow(a)="; "*"; noshow(nocount); "*"; aaa, nocount
 '       tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
        End If
        GoTo line_29190
    End If
    nocount = nocount + 1
    noshow(nocount) = Left(temps, tt - 1)
    temps = Right(temps, Len(temps) - tt)
'        Print "noshow(b)="; "*"; noshow(nocount); "*"; aaa, nocount
'        tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
    GoTo line_29150
line_29190:
    save_line = "29190"
    For II = 1 To 10
    If Len(noshow(II)) < 3 Then
        GoTo line_29199
    End If
        temps = noshow(II) + ""
        tt = InStr(1, " ", temps)
    save_line = "29190a"
        If tt > 1 Then
            ttt1 = Len(temps) - tt
            tempss = Left(temps, ttt1)
            noshow(II) = tempss + ""
            'strip off any trailing spaces
        End If
line_29199:
    Next II
    For II = 1 To nocount
'        noshow(II) = " " + UCase(noshow(II)) + " "
        noshow(II) = UCase(noshow(II))
  '      Print "noshow(II)="; "*"; noshow(II); "*"; aaa, II
  '      tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
        'put 1 space before and 1 after only
    Next II
Return
line_29200:     'screen saver logic here
    screencount = 0
    save_line = "29200"
    For II = 1 To 10
        screensave(II) = ""
    Next II
    If Len(Cmd(26)) < 3 Then
        GoTo line_29290
    End If
'    nocount = 0
    temps = Cmd(26) + ""
line_29250:
    save_line = "29250"
    tt = InStr(temps, " ")
    If tt = 1 Then
        temps = Right(temps, Len(temps) - 1)
        GoTo line_29250 'strip out any leading spaces
    End If
    If tt = 0 Then
        If Len(temps) > 2 Then
            screencount = screencount + 1
            screensave(screencount) = temps
'        Print "screensave(a)="; "*"; screensave(screencount); "*"; aaa; screencount
'        tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
        End If
        GoTo line_29290
    End If
    screencount = screencount + 1
    screensave(screencount) = Left(temps, tt - 1)
    temps = Right(temps, Len(temps) - tt)
    GoTo line_29250
line_29290:
    save_line = "29290"
    For II = 1 To 10
    If Len(screensave(II)) < 3 Then
        GoTo line_29299
    End If
        temps = screensave(II) + ""
        tt = InStr(1, " ", temps)
    save_line = "29290a"
        If tt > 1 Then
            ttt1 = Len(temps) - tt
            tempss = Left(temps, ttt1)
            screensave(II) = tempss + ""
            'strip off any trailing spaces
        End If
line_29299:
    Next II
    For II = 1 To screencount
'        screensave(II) = " " + UCase(screensave(II)) + " "
        screensave(II) = UCase(screensave(II))
        'put 1 space before and 1 after only
    Next II

Return

line_29300:
        changes = "NO"      'december 2 2000
    For tt = 1 To criptcnt
 '           frmproj2.Caption = " testing 444 *" + cript1(tt) + "*" + CStr(cript3(tt)) + " " + CStr(criptcnt) + " " + CStr(tt) '27 October 2004
 '           new_delay_sec = 0.1      '27 October 2004  testing only
 '           GoSub line_30300        '27 October 2004  testing only
    III = 1
 '       in_str = cript1(tt)
 '       out_str = cript2(tt)
'        ttt1 = Len(in_str)
'        II = Len(out_str)
'        xtemp = Len(cript2(tt))
        save_line = "29310"     '17Dec2017
'            frmproj2.Caption = " testing 29310 *" + Len(Trim(aaa)) '17Dec2017
'17Dec2017 this line vvv is a problem error trapped at line 29310
        JJ = Len(aaa)
    If case_yes = "Y" And Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then   '27 October 2004
        If aaa = cript1(tt) Then
'            frmproj2.Caption = " testing 444a " + CStr(cript3(tt)) + " " + CStr(criptcnt) '27 October 2004
'            new_delay_sec = 1       '27 October 2004  testing only
'            GoSub line_30300        '27 October 2004  testing only
            temp2 = temp2 + 1       'count of changes done
            aaa = " "
            cript1(tt) = "xyxyxyxyxyxyxyxyxy"           '27 October 2004 clear it
            GoTo line_29399                    '29 October 2004
        End If                                          '27 October 2004
    End If                                              '27 October 2004
    If case_yes = "Y" And InStr(aaa, cript1(tt)) = 0 Then
'        Print #FileFile, aaa
        GoTo line_29399
    End If
    If case_yes = "N" And Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then   '27 October 2004
        If UCase(aaa) = cript1(tt) Then
'            frmproj2.Caption = " testing 444aa " + CStr(cript3(tt)) + " " + CStr(criptcnt) '27 October 2004
'            new_delay_sec = 1       '27 October 2004  testing only
'            GoSub line_30300        '27 October 2004  testing only
        temp2 = temp2 + 1       'count of changes done
            aaa = " "
            cript1(tt) = "xyxyxyxyxyxyxyxyxy"           '27 October 2004 clear it
            GoTo line_29399                    '29 October 2004
        End If                  '27 October 2004
    End If                                              '27 October 2004
    If case_yes = "N" And InStr(UCase(aaa), cript1(tt)) = 0 Then
'        Print #FileFile, aaa
        GoTo line_29399
    End If

line_29320:
    If case_yes = "Y" Then
        II = InStr(III, aaa, cript1(tt))
    Else
        II = InStr(III, UCase(aaa), cript1(tt))
    End If
    If II = 0 Then
'        Print #FileFile, aaa
        GoTo line_29399
    End If
    III = II + 1
 '27 October 2004 here do some checking
    aaa = Left(aaa, II - 1) + cript2(tt) + Right(aaa, JJ + 1 - II - cript3(tt))
    JJ = Len(aaa)       'november 13 2000
'        Print aaa
'        Print tt, JJ, cript3(tt), II
'        tt1 = InputBox("testing", , , 4400, 4500)  'TESTING ONLY
'    If UCase(tt1) = "X" Then
'        GoTo line_29090
'    End If
    temp2 = temp2 + 1       'count of changes done
    changes = "YES"         'december 2 2000
    GoTo line_29320

line_29399:
    Next tt             '17Dec2017 comment end of loop
        If changes = "YES" And crlf = "NO" Then
        Print #FileFile, aaa;
    Else
    If Left(UCase(Cmd(71)), 11) = "BATCHFILE==" Then   '03 November 2004
        If Len(Trim(aaa)) < 1 Then
        Else
            Print #FileFile, aaa
        End If              '03 November 2004
    Else
        Print #FileFile, aaa
    
    End If                  '03 November 2004 keep all the blank lines from being put to output file
    End If              'do not print crlf if sss requested
                        'instead of rrr ie for the From:
                        'change december 2 2000
    
    Return

line_29400:             'november 19 2000
'   For tt = 1 To criptcnt
'    cript3(tt) = Len(cript1(tt))
'    Next tt
'    save_line = "29400"
    tt1 = ""        'the output string after translation
'17Dec2017 same error here...
    
    save_line = "29400"        '17dec2017
    
    JJ = Len(aaa)
line_29410:
    FFF = ""        'indicate if search string found.
    xtemp = Left(aaa, 1)
    For tt = 1 To criptcnt
'        in_str = cript1(tt)
'        out_str = cript2(tt)
'        ttt1 = Len(in_str)
'    If cript3(tt) > JJ Then        'november 23 2000
'        GoTo line_29450   'more characters than in string november 23 2000
'    End If                         'november 23 2000
'    If case_yes = "N" Then GoTo line_29430
'    If cript3(tt) = 1 And xtemp <> cript1(tt) Then GoTo line_29450 'november 23 2000
'    If Left(aaa, cript3(tt)) <> cript1(tt) Then november 23 2000 see following line
'    If Left(aaa, 1) <> cript1(tt) Then
    If xtemp <> cript1(tt) Then
        GoTo line_29450
    End If
'    GoTo line_29440
'line_29430:
'    If cript3(tt) = 1 And UCase(xtemp) <> UCase(cript1(tt)) Then GoTo line_29450
'    If UCase(Left(aaa, cript3(tt))) <> UCase(cript1(tt)) Then
'        GoTo line_29450
'    End If
line_29440:
'    aaa = Right(aaa, JJ - cript3(tt))  'november 23 2000
    aaa = Right(aaa, JJ - 1)
    FFF = "Y"
    GoTo line_29460
line_29450:
    
    Next tt

line_29460:
    If FFF <> "Y" Then
        tt1 = tt1 + Left(aaa, 1)
        If JJ > 1 Then
            aaa = Right(aaa, JJ - 1)
        Else
            aaa = ""
        End If
    Else
        temp2 = temp2 + 1
        tt1 = tt1 + cript2(tt)
    End If
'    JJ = Len(aaa)      'november 23 2000
    JJ = JJ - 1
    If JJ > 0 Then
        GoTo line_29410
    End If
    
    If encript <> "MYSTUF" Then
        Print #FileFile, tt1
    Else
        aaa = tt1
    End If
    Return

line_29500:
    save_line = "29500"
    FileFile = FreeFile
    Open FileExt For Input As #FileFile
    temp2 = 0
    criptcnt = 0
    save_line = "29510"
line_29510:
    Line Input #FileFile, aaa
    temp2 = temp2 + 1
    Line Input #FileFile, tt1
    temp2 = temp2 + 1
    criptcnt = criptcnt + 1
    If encript = "CRIPT" Then
        cript1(criptcnt) = aaa
        cript2(criptcnt) = tt1
    Else
        cript1(criptcnt) = tt1
        cript2(criptcnt) = aaa
    End If
    GoTo line_29510
line_29520:
    save_line = "29520"
    Close #FileFile
    For tt = 1 To criptcnt
        cript3(tt) = Len(cript1(tt))
    Next tt

    'tempss = InputBox("cript.txt read" + CStr(criptcnt), , , 4400, 4500) 'TESTING ONLY
    Return          'november 20 2000 read in the encription file info.

    
'timer subroutine set delay_sec to seconds to hold wait sleep
line_30000:
    DoEvents
    
'january 23 2002    If delay_sec < 1 Then
'    tempss = InputBox("testing delay" + Format(delay_sec, "###0.000"), , , 4400, 4500) 'january 23 2002
'04 february 2003 comment out below want to set WAIT=0 for avi file displays
'    If delay_sec < 0.0001 Then delay_sec = Val(Cmd(27)) '02 may 2002
    If delay_sec < 0.00001 Then
        GoTo line_30020
    End If
            
        
    '29 September 2003 add or subtract amount in cmd() to delay_sec
    If adjust_sec <> 0 Then
        delay_sec = delay_sec + adjust_sec
    End If                      '29 September 2003
    temp_cnt1 = 0       '12 April 2004
line_30005:
            
'12 April 2004
    If slomo = False And mpg_file = "YES" Then
            i = mciSendString("status video1 position", vs, 255, 0)
            temp3 = InStr(vs, Chr$(0)) '19 March 2004
            temp_cnt = Val(Left(vs, temp3 - 1)) '19 March 2004
'12 April 2004        If line_start_point + (delay_sec * 1000) >= temp_cnt + 0.01 Then
'14 April 2004 this is where a problem cropped up (just adding 0.1 fixed it??)
'14 April 2004        If line_start_point + (delay_sec * 1000) >= temp_cnt + 0.01 And temp_cnt  > temp_cnt1 Then
'        If line_start_point + (delay_sec * 1000) >= temp_cnt + 0.01 And temp_cnt + 0.1 > temp_cnt1 Then
        If line_start_point + (delay_sec * 1000) >= temp_cnt Then      '23Jan2018 the end of the drag race videos were not showing fixed
            DoEvents        '10 April 2004 allow for interrupt in full speed video....
            temp_cnt1 = temp_cnt '12 April 2004
            GoTo line_30005
        End If
        i = mciSendString("pause video1", 0&, 0, 0)
        mpg_file = "NO"         '24 March 2004
        DoEvents
'            frmproj2.Caption = "test 24 March 2004  " + CStr(line_start_point) + " " + CStr(delay_sec * 1000) + " " + CStr(temp_cnt) '24 March 2004
'    tempss = InputBox("testing delay" + Format(delay_sec, "###0.000"), , , 4400, 4500) 'january 23 2002
        GoTo line_30020         '24 March 2004...
    End If
        
line_30012:
'    If slomo <> True Then   '08 January 2004 so when enter hit it will be in the pause routine
'        DoEvents                '03 november 2002 turn off when debugging problems etc
'        DoEvents                '23 September 2003 add another just for more sharing...
'    End If                  '08 January 2004 just the if statement around the doevents added
    
'24 March 2004 2004======================================
'15Nov2011    If slomo Then                   'the if slomo = true loop here ***vip***slomo ***hint***
'15Nov2011 this made it work for a again and full speed otherwise it was playing slowmo 2nd time
    If slomo And play_speed <> 1000 Then                   'the if slomo = true loop here ***vip***slomo

line_30012a:
            i = mciSendString("status video1 position", vs, 255, 0)
            DoEvents            '26Apr2012might want to skip this one because background jobs could hog and cause extended play
            temp3 = InStr(vs, Chr$(0)) '10 April 2004
'            temp_cnt = Val(Left(vs, temp3 - 1)) '10 April 2004
            temp_cnt = Val(vs) '
            
'10 April 2004            temp_cnt = Val(vs)
            If temp_cnt < last_vs + pad_time Then GoTo line_30012a
            
            tempdata = CStr(temp_cnt)
            last_vs = Val(vs)

        If replay_yn = True Then
            If temp_cnt >= replay_pos Then
'                tempss = InputBox("testing replay position " + CStr(temp_cnt), , , 4400, 4500) '10 April 2004
                play_speed = hold_speed '10 April 2004
                replay_yn = False       '10 April 2004
            End If
        End If                      '10 April 2004
           
        If UCase(Left(Cmd(63), 8)) = "SHOWELAP" Then
            DoEvents
'            frmproj2.Caption = LCase(temptemp) + " (Replay by Spectate Swamp) " + tempdata + " of " + CStr(video_length) '18 February 2004
'12Mar2012            frmproj2.Caption = LCase(temptemp) + " (by Spectate Swamp) " + tempdata + " of " + CStr(video_length) '18 February 2004
            frmproj2.Caption = LCase(temptemp) + " elapsed** " + tempdata + " of " + CStr(video_length) '18 February 2004
            DoEvents
            tempdata = ""
        End If
        If begin = "YES" And line_start_point > temp_cnt Then GoTo line_30016   '15Sep2011
        If begin_point <> 0 And temp_cnt < line_start_point Then GoTo line_30016 '20Sep2011
        If resume_str = "NO" Then GoTo line_30016                   '25Sep2011   ==
 'every once in a while it would not do the resume so I added the -50 below
'        If resume_str = "YES" And temp_cnt >= line_start_point + delay_sec * 1000 - 50 Then
        If resume_str = "YES" And temp_cnt >= line_start_point + delay_sec * 1000 Then  '23Jan2018
                play_speed = 1000              '25Sep2011
                save_play_speed = 1000          '25Sep2011
'23Jan2018                delay_sec = 1000                   '25Sep2011
                resume_str = "NO"                   '25Sep2011
                GoTo line_30016 '25Sep2011
        End If          '25Sep2011
        i = mciSendString("pause video1", 0&, 0, 0)
        DoEvents
        new_delay_sec = (1000 / play_speed) * (pad_time / 1000)
'            frmproj2.Caption = "test 24 March 2004  " + CStr(new_delay_sec) '24 February 2004
        
        motion_in = Timer       '01 March 2004 changed from slomo_in etc
line_30012b:
'try and slow the delay for character pause. it would not do less than 1/10 sec or so?
    If new_delay_sec > 0.02 Then     '05 February 2008
          DoEvents
        motion_out = Timer
        If motion_out + 43200 < motion_in Then
            motion_in = motion_in - 86400
        End If
        If motion_out - motion_in < new_delay_sec Then
            GoTo line_30012b
        End If
    Else                               ' 05 February 2008
        temp_cnt1 = Int(new_delay_sec * 100000)
        For temp_cnt = 1 To temp_cnt1
            For II = 1 To 10
                DoEvents
            Next II
        Next temp_cnt
    
    End If                              '05 February 2008
        
'26 March 2004 try and back it up a bit before resume. that should make it even slower...
'         i = mciSendString("seek video1 " + CStr(last_vs - pad_time), 0&, 0, 0) '26 March 2004
         i = mciSendString("resume video1", 0&, 0, 0)
line_30016:                 '15Sep2011
        DoEvents
    End If                  'end of "If slomo = true then ***vip***slomo ***hint***
    
If resume_str = "YES" And temp_cnt >= line_start_point + delay_sec * 1000 Then slomo = False  '25Sep2011
'If again_str = "DONE" And temp_cnt < line_start_point + delay_sec * 1000 Then slomo = False   '14Nov2011
    DoEvents        '09 March 2004
'    slomo = True    '19 March 2004  just so that full speed ie=1000 will get the elapsed displayed
'        slomo = True '14 April 2004 testing only
        'handy for testing too
        
    If slomo Then
        DoEvents        '16 March 2004
        i = mciSendString("status video1 position", vs, 255, 0) '04 March 2004
        temp_double = Val(Trim(vs))         '18 March 2004
        tempdata = CStr(temp_double)             '17 March 2004
        If temp_double < 1 Then
'14Jan2012            xtemp = InputBox(" test 18 March 2004 " + CStr(line_start_point + (delay_sec * 1000)) + " > " + Trim(vs), , xx1 - offset1, yy1 - offset2)
        End If                  '18 March 2004
        
        DoEvents
        
'13 April 2004 need to check temp_cnt against video_length here and ship out meebee
        If temp_double >= video_length - 100 Then
            GoTo line_30020
        End If                  '13 April 2004
        
        If line_start_point + (delay_sec * 1000) <= Val(vs) + 0.01 Then
'        xtemp = InputBox(" test 24 March 2004 " + CStr(line_start_point) + "  " + CStr((line_delay_sec * 1000)) + " <= " + Trim(vs), , xx1 - offset1, yy1 - offset2)   'test
            GoTo line_30020
        Else
            GoTo line_30012
        End If
    End If              '04 March 2004

' 14 January 2004 test show the play_speed here???
'        xtemp = InputBox(" test 14 january 2004 " + CStr(play_speed), "testing Prompt   ", , xx1 - offset1, yy1 - offset2) 'test
 
    GoTo line_30020 'january 23 2002 ie just skip the logic below every time
line_30018:         'february 08 2002 deactivated as above
    ampm = ""
    temps = Format(Now, "ddddd ttttt") 'start timer time
'    Print "start timeer="; temps
    temp1 = InStr(2, temps, " ")
    temps = Right(temps, Len(temps) - temp1) 'just the hrs and minutes
    temp1 = InStr(2, temps, "M")    'check if US format
    If temp1 <> 0 Then
'        Print tt, JJ, ttt1
 '       tt1 = InputBox("PM AM found", , , 4400, 4500)  'TESTING ONLY
        temp2 = InStr(2, temps, "PM")
        If temp2 <> 0 Then
            ampm = "PM"
        End If
        If temp2 = 0 Then
            ampm = "AM"
        End If
        temps = Left(temps, Len(temps) - 3) 'strip of AM PM
    End If
    temp1 = InStr(1, temps, ":")
    hhour = Left(temps, temp1 - 1)
    temps = Right(temps, Len(temps) - temp1)
    temp1 = InStr(1, temps, ":")
    mminute = Left(temps, temp1 - 1)
    ssecond = Right(temps, Len(temps) - temp1)
    hhhour = Val(hhour)
    mmminute = Val(mminute)
    sssecond = Val(ssecond)
    If ampm = "PM" Then
        hhhour = hhhour + 12
    End If
    temp11 = hhhour * 3600 + mmminute * 60 + sssecond  'moved from below march 14 2001

line_30010:
    DoEvents            '03 november 2002 turn off when debugging (other jobs on computer ??)
    save_line = "30010"
    ampm = ""
    temps = Format(Now, "ddddd ttttt") 'current time
    temp1 = InStr(2, temps, " ")
    temps = Right(temps, Len(temps) - temp1) 'just the hrs and minutes
    temp1 = InStr(2, temps, "M")    'check if US format
    If temp1 <> 0 Then
        temp2 = InStr(2, temps, "PM")
        If temp2 <> 0 Then
            ampm = "PM"
        End If
        If temp2 = 0 Then
            ampm = "AM"
        End If
        temps = Left(temps, Len(temps) - 3) 'strip of AM PM
    End If
    temp1 = InStr(1, temps, ":")
    hhour = Left(temps, temp1 - 1)
    temps = Right(temps, Len(temps) - temp1)
    temp1 = InStr(1, temps, ":")
    mminute = Left(temps, temp1 - 1)
    ssecond = Right(temps, Len(temps) - temp1)
    chhour = Val(hhour)
    cmminute = Val(mminute)
    cssecond = Val(ssecond)
    If ampm = "PM" Then
        chhour = chhour + 12
    End If
    If chhour < hhhour Then
        chhour = chhour + 23
    End If


'    If cmminute < mmminute Then
'march 14 2001        cmminute = cmminute + 59
'    End If
'    If cssecond < sssecond Then
'march 14 2001        cssecond = cssecond + 59
'    End If
'    temp1 = hhhour * 3600 + mmminute * 60 + sssecond  'could move this out of the loop?
    temp2 = chhour * 3600 + cmminute * 60 + cssecond
'    If temp2 < temp1 + delay_sec Then      'march 14 2001
    If temp2 < temp11 + delay_sec Then
        GoTo line_30010
    End If
     '   Set Picture = LoadPicture() 'clear any picture
line_30020:
'17Dec2017 **hint** budge needs similar logic as the again but just again and again....
'            frmproj2.Caption = LCase(temptemp) + " budging***** " + budge_str + " " + tempdata + " of " + CStr(video_length) '06Jan2018
'11Jan2018        budge_str = "YES"       '07Jan2018 test only
        totbud = 0          '11Jan2018
        If budge_str = "YES" Then       '06Jan2018

line_30022:
            i = mciSendString("pause video1", 0&, 0, 0)         '07Jan2018
            
            totbud = totbud + 1
            If totbud >= maxbud Then GoTo line_30023

'            frmproj2.Caption = LCase(temptemp) + " budging***** " + budge_str + " " + tempdata + " of " + CStr(video_length) '06Jan2018
'testtest = "resume_str=" + resume_str + " motion_yn=" + motion_yn + " start_point=" + CStr(start_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " keep_line_delay_sec=" + CStr(keep_line_delay_sec) + " " '08Nov2011 teststring
'       testtest = InputBox("again info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'            temp_cnt = 0                '08Nov2011
'28Dec2011  these 2 caused trouble          line_start_point = keep_line_start_point     '08Nov2011
'28Dec2011  these 2 caused trouble          start_point = keep_line_start_point     '08Nov2011 14nov2011 remove comment out
'25Feb2012 the problem was that wait=?? needed to be added in the catalog function whoo hoo
'        start_point = keep_line_start_point     '25Feb2012
'15Dec2011            start_point = 10                        '08Nov2011 14Nov2011 remove comment  out
            tempdata = String(totbud, "*")      '09Jan2018
            frmproj2.Caption = LCase(temptemp) + " budging== " + tempdata + " " + budge_str + " " + CStr(begin_point) + " of " + CStr(video_length) + " " + tempdata '06Jean2018
'        temptemp = InputBox("07 January 2018 test ", "budge " + CStr(totbud), , xx1 - 5000, yy1 - 5000)     '07Jan2018
            last_vs = 0                             '08Nov2011? this was one of the problems yahoo
'            begin = keep_begin_point + (totbud * smlbud)                     '08Nov2011
'15Dec2011 mess with these settings it does not see the start== point somehow on the 2nd one  (wait=?? fixed it)
'            begin_point = keep_line_start_point + (totbud * smlbud)           '07Jan2018
            begin_point = keep_begin_point + (totbud * smlbud)           '07Jan2018
            resume_str = keep_resume_str              '08Nov2011
            play_speed = keep_play_speed             '08Nov2011
'            play_speed = save_play_speed            '14Nov2011
            keep_slomo = True                       '14Nov2011 force for normal speed without it craps??
'            keep_play_speed = 1000                  '14Nov2011 force this to test
            save_play_speed = keep_play_speed       '08Nov2011
            slomo = keep_slomo                     '08Nov2011
            line_delay_sec = keep_line_delay_sec    '08Nov2011
            delay_sec = keep_line_delay_sec          '08Nov2011
'            delay_sec = 3          '25Feb2012   testing set it on the line if it works
'           i = mciSendString("play video1 from " + CStr(begin_point), 0&, 0&, 0&)
            frmproj2.Caption = LCase(temptemp) + " budging******** " + tempdata + " of " + CStr(video_length) '06Jan2018
           i = mciSendString("play video1 from " + CStr(begin_point), 0&, 100, 100)
'08Nov2011 not having the " wait" allows for CR to interrupt the video otherwise it plays straight through
'            i = mciSendString("play video1 from " + CStr(begin_point) + " wait", 0&, 0, 0) 'this is where the display happens
            GoTo line_30012         '08Nov2011
line_30023:
'            i = mciSendString("pause video1", 0&, 0, 0)         '11Jan2018
            frmproj2.Caption = LCase(temptemp) + " budging***** " + tempdata + " of " + CStr(video_length) '06Jan2018
'        temptemp = InputBox("07 January 2018 test ", "budge " + CStr(totbud), , xx1 - 5000, yy1 - 5000)     '07Jan2018
            totbud = 0          '06Jan2018
            GoTo line_30012    '06Jan2018
        End If              '08Nov2011
        
        If again_str = "YES" Then       '08Nov2011
'testtest = "resume_str=" + resume_str + " motion_yn=" + motion_yn + " start_point=" + CStr(start_point) + " keep_line_start_point=" + CStr(keep_line_start_point) + " begin_point=" + CStr(begin_point) + " play_speed=" + CStr(play_speed) + " slomo=" + CStr(slomo) + " keep_line_delay_sec=" + CStr(keep_line_delay_sec) + " " '08Nov2011 teststring
'       testtest = InputBox("again info " + testtest, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            again_str = "DONE"
'            temp_cnt = 0                '08Nov2011
'28Dec2011  these 2 caused trouble          line_start_point = keep_line_start_point     '08Nov2011
'28Dec2011  these 2 caused trouble          start_point = keep_line_start_point     '08Nov2011 14nov2011 remove comment out
'25Feb2012 the problem was that wait=?? needed to be added in the catalog function whoo hoo
'        start_point = keep_line_start_point     '25Feb2012
'15Dec2011            start_point = 10                        '08Nov2011 14Nov2011 remove comment  out
            last_vs = 0                             '08Nov2011? this was one of the problems yahoo
            begin = keep_begin                      '08Nov2011
'15Dec2011 mess with these settings it does not see the start== point somehow on the 2nd one  (wait=?? fixed it)
            begin_point = keep_begin_point          '08Nov2011
            resume_str = keep_resume_str            '08Nov2011
            play_speed = keep_play_speed            '08Nov2011
'            play_speed = save_play_speed            '14Nov2011
            keep_slomo = True                       '14Nov2011 force for normal speed without it craps??
'            keep_play_speed = 1000                  '14Nov2011 force this to test
            save_play_speed = keep_play_speed       '08Nov2011
            slomo = keep_slomo                     '08Nov2011
            line_delay_sec = keep_line_delay_sec    '08Nov2011
            delay_sec = keep_line_delay_sec          '08Nov2011
'            delay_sec = 3          '25Feb2012   testing set it on the line if it works
           i = mciSendString("play video1 from " + CStr(begin_point), 0&, 0&, 0&)
'08Nov2011 not having the " wait" allows for CR to interrupt the video otherwise it plays straight through
'            i = mciSendString("play video1 from " + CStr(begin_point) + " wait", 0&, 0, 0) 'this is where the display happens
            GoTo line_30012         '08Nov2011
        End If              '08Nov2011
'16 August 2004
    If motion_yn <> "YES" Then   '16 August 2004
        new_delay_sec = delay_sec   '16 August 2004
        GoSub line_30300        '16 August 2004
    End If                  '16 August 2004
'    frmproj2.Caption = "delay_sec=" + CStr(delay_sec) '16 August 2004
    last_vs = 0
    delay_sec = Val(Cmd(27))    '25 July 2003
    DoEvents                '03 november 2002 deactivate when in debug mode for more control??
    
    Return
    
Return

line_30300:     'mini timer routine 26 February 2004
    frmproj2.Caption = " save_line " + save_line + " delay_sec=" + CStr(delay_sec) '03mar2017
    
        motion_in = Timer       '01 March 2004 changed from slomo_in etc
line_30316:
'try and slow the delay for character pause. it would not do less than 1/10 sec or so?
    If new_delay_sec > 0.02 Then     '05 February 2008
          DoEvents
        motion_out = Timer
        If motion_out + 43200 < motion_in Then
            motion_in = motion_in - 86400
        End If
        If motion_out - motion_in < new_delay_sec Then
            GoTo line_30316
        End If
    Else                               ' 05 February 2008
        temp_cnt1 = Int(new_delay_sec * 100000)
        For temp_cnt = 1 To temp_cnt1
            For II = 1 To 10
                DoEvents
            Next II
        Next temp_cnt
    
    End If                              '05 February 2008
    
Return                                  '26 February 2004

'     W R I T E  C O N T R O L . T X T  H E R E   **hint**
'control.txt control control.txt control.txt control.txt control control file here
'create control file open for output etc
'write to the control file here
line_30500:     'november 03 2000 control.txt and control1.txt file create print
    save_line = "30500"
    If visual_impared <> "YES" Then
        Print #FileFile, "34"    '     lines per page   1
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "11"    '     lines per page   1
    End If
    If visual_impared <> "YES" Then
        Print #FileFile, "12"    '     font size        2
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "36"    '     font size        2
    End If
    If visual_impared <> "YES" Then
        Print #FileFile, "3"    ' aqua     background color 3
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "6"   ' purple 3   background color 6=brown or gold
    End If
    If visual_impared <> "YES" Then
        Print #FileFile, "12"   ' white     text color       4
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "0"    ' black    text color       4
    End If
'midway4 thru the program appx'==================================================================================
    
        Print #FileFile, "12"    '     Hi-lite color    5
        Print #FileFile, "/"     '     seperator chr    6
        Print #FileFile, "dateno"  '  or dateyes        7
        Print #FileFile, "c:\windows\explorer.exe" 'jpg 8  most have this file...
'        Print #FileFile, "c:\winnt\explorer.exe" 'jpg 8
'17 october 2002 made the change below to prevent system from crashing without explorer
'see "photo 6b-1" display
'         Print #FileFile, "c:\program files\accessories\wordpad.exe" 'jpg 8 17 october 2002
        Print #FileFile, "c:\PROGRAM FILES\ACCESSORIES\MSPAINT.EXE"  '  bmp disp  9
        Print #FileFile, "c:\PROGRAM FILES\ACCESSORIES\WORDPAD.EXE"  '  text chg 10
'        Print #FileFile, "c:\PROGRAM FILES\windows nt\ACCESSORIES\WORDPAD.EXE"  '  text chg 10
'        Print #FileFile, "c:\files.txt"    ' Files     11
'        Print #FileFile, "c:\search.txt"   '  search   12
'        Print #FileFile, "c:\temp.txt"     'displayed  13
        Print #FileFile, "files.txt"    ' Files     11 december 3 2000
        Print #FileFile, "search.txt"   '  search   12 december 3 2000
        Print #FileFile, "10"    'over_lap on line_len january 10 2001 13
        Print #FileFile, "enterdateyes" 'start date on 14
        Print #FileFile, "exitdateno"  'end date on   15
        Print #FileFile, "C:\PROGRAM FILES\Microsoft Office\Office\MSPUB.EXE"      'default drive 3    16
        Print #FileFile, "13000"    'xx1 P1 prompt      17
        Print #FileFile, "10700"    'yy1 P1 prompt      18
'        Print #FileFile, "c:\replace.txt" 'search replace 19
        Print #FileFile, "replace.txt" 'search replace 19 december 3 2000
        Print #FileFile, "WW"       'option default    20 see item cmd(42)
    If visual_impared <> "YES" Then
        Print #FileFile, "175"       'wrap line length  21
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "55"       'wrap line length  21
    End If
    If visual_impared <> "YES" Then
        Print #FileFile, "16"       'context lines     22
    End If
    
    If visual_impared = "YES" Then
        Print #FileFile, "3"       'context lines     22
    End If
    'when adding new elements to the control.txt file add the detail notes here comment it well
        Print #FileFile, "d:\search\tempfold\"       'photo copy directory   23 march 17 2001
        Print #FileFile, "xxx.c:\family1\scn000"    'vvv paste string on entry 24
        Print #FileFile, " photo 1980 "    'ggg paste string on entry 25
        Print #FileFile, "photo"    'search string for screen saver 26
        Print #FileFile, "3"       'default sleep time for screen saver 27
        Print #FileFile, "noshow mpotj"    ' no show element skip pics with these 28
        Print #FileFile, "12"    ' 29 secondary display color line green 10 (12 = red)
        Print #FileFile, "Y"    ' 30 close out form Y/N
        Print #FileFile, "xxx."    ' 31 if this exists hilite it hilite_this
'        Print #FileFile, "c:\cript.txt"    'encription control file 32
        Print #FileFile, "cript.txt"    'encription control file 32 december 3 2000
        Print #FileFile, "PHOTO"    ' default search string on start 33 was "D" before
        Print #FileFile, "5"    ' context extract window in lines of display january 01 2001 34
        Print #FileFile, "c:\"    ' 35 the input directory for the gf option
'19 September 2004        Print #FileFile, "ALL"    ' 36 jpg bmp tif etc etc
        Print #FileFile, "MPG"    ' 36 jpg bmp tif etc etc
        Print #FileFile, "MPEG_VIDEOS.txt"    ' 37 using gf to get files they are catalogued here
        Print #FileFile, "nostretch"    ' 38 stretch or normal display p1 or p2
        Print #FileFile, "noautoredraw"    ' 39 allows for screen refresh between sessions
'07Aug2012        Print #FileFile, "A:\ B:\ C:\ D:\ E:\ F:\ G:\ H:\ I:\"    ' 40 Cmd(40) allow for program to run from a non writeable cd (if app.path 17Sep2010 add everything besides a: and c: clear to disable cd and dvd updates ***vip***
'07Aug2012 default to non update for most devices c is ok but this is just simpler
        Print #FileFile, "A:\ B:\ C:\ D:\"    ' 40 Cmd(40) allow for program to run from a non writeable cd (if app.path 17Sep2010 add everything besides a: and c: clear to disable cd and dvd updates ***vip***
                             'not in list then don't update files or search or control files
                             '20 july 2002
        Print #FileFile, "WW"       '41 cmd(41) if photo file default to this as option prompt#2
        Print #FileFile, "C"        '42 cmd(42) if text file default to this as option prompt #2
        
        Print #FileFile, "c:\winnt\explorer.exe" '43    see #8 above
        Print #FileFile, "c:\PROGRAM FILES\windows nt\ACCESSORIES\WORDPAD.EXE"  '  44 see #10 above
    If visual_impared <> "YES" Then
        Print #FileFile, "SEQUENTIAL RANDOM MANUAL"       ' 45  AUTO RUN PROGRAM CHECK MULTIPLE NAMES ALLOWED HERE
    End If
    If visual_impared = "YES" Then
        Print #FileFile, "SEQUENTIAL"       ' 45  AUTO RUN PROGRAM CHECK MULTIPLE NAMES ALLOWED HERE
    End If
'26Jun2012        Print #FileFile, "SEQUENTIAL RANDOM MANUAL"    ' 45  AUTO RUN PROGRAM CHECK MULTIPLE NAMES ALLOWED HERE
        Print #FileFile, "MPEG_VIDEOS.TXT"    ' 46    FIRST PROMPT FILE SELECT IE #1 FILE
        Print #FileFile, "WW"    ' 47 SEARCH OPTIONS allow for "TT2.5 RAND P2 WW"
        Print #FileFile, "PHOTO"    ' 48    SEARCH STRING FOR AUTO RUN
        Print #FileFile, "NORANDOM"    ' 49 RANDOM OR NORANDOM ON THE PHOTO SEARCH 09 december 2002
        Print #FileFile, "SHOWFILES"    ' 50 ON MERGED FILES SHOW OR NOT THE FILE NAME IN THE CAPTION AREA
        Print #FileFile, "1084"    ' 51 avi file size x
        Print #FileFile, "800"    ' 52  avi file size y
        Print #FileFile, "SHOWVIDEO"    ' 53 set to SHOWVIDEO if avi installed 10 february 2003
        Print #FileFile, "SHOWDUPS"    ' 54   28 FEBRUARY 2003
'vvversion make changes here to add to the control files...
        Print #FileFile, "00"    ' 55 Context Clear line count 18 march 2003 ver=1.01
        Print #FileFile, "PHOTO_DETAIL"    ' 56 Display photo detail info on pictures must be normal not stretch ie image control element
                                            '11Feb2017 default to photo detail for overprint of pictures with text
                                            'ver=1.02 above photo_detail
        Print #FileFile, "noSHOWTIME"         ' 57 ver=1.03 time display on off switch
        Print #FileFile, "DEFAULT_TO_CD"    ' 58 ver=1.06 if set then default to cd over-ride drive # 21Sep2010 change from noDEFAULT_TO_CD
        Print #FileFile, "0.00"            ' 59 Adjust wait time by this value
        Print #FileFile, "0.00"    ' 60 freeze_sec if set for mpeg files only
        Print #FileFile, "SETSPEED1000"    ' 61 to allow for switching set speed option off 30 October 2003 1000 = full speed 500=1/2
        Print #FileFile, "0.00668"    ' 62 slow motion segment length (with 0.01667 play for 2 hundredths)
        Print #FileFile, "SHOWELAP"    ' 63 show the elapsed time of a video (slomo only). 11 February 2004
        Print #FileFile, "0"    ' 64  21 March 2004 pad / add to elapsed amount 30 50 100 etc
        Print #FileFile, "noRANDBEG"    ' 65 23 March 2004 randomly generate a begin point for mpg play
        Print #FileFile, "125"    ' 66 10 April 2004 the speed for the replay of a paused video
        Print #FileFile, "noTHUMB"    ' 67 14 April 2004 have as a control (handy to run CD's with rand & thumb)
        Print #FileFile, "noEOF_STOP"    ' 68 12 September 2004 (WW)allow for backgrd jobs to halt after 1 pass cmd(68)
        Print #FileFile, "noHIT_STOP"    ' 69 17 September 2004 (PP)allow for backgrd jobs to halt after first match
        Print #FileFile, "noFOREGROUND"  ' 70 17 September 2004   have the second job get focus ie Foreground & hi-lited
        Print #FileFile, "noBATCHFILE==BATCHFILE.TXT"    ' 71 27 October 2004 allow for batch input especially for rrr search and replace
        Print #FileFile, "noRESULTS.TXT"    ' 72 08 November 2004 save a play list for video, pics, music handy to check previously played segment.
        Print #FileFile, "noFILESWITCH"    ' 73 22 November 2004 when control file switches switch the file too?
        Print #FileFile, "noEOFCMD==CHGFILE.TXT"    ' 74 26 November 2004 on end of file go to this file instead of loop
        Print #FileFile, "VIDEOSTOP"    ' 75 12 December 2004 used for dvd and cds do not allow interrupt use the wait command
        Print #FileFile, "noLINEPAUSE==0.65"    ' 76 24 December 2004 on each print line allow for a delay per line (like scrolling text)
        Print #FileFile, "noCHARACTERPAUSE==.0333"    ' 77 03 January 2005 allow for characters to be displayed 1 at a time if easy
        Print #FileFile, "HTM, FRM, CLS, "    ' 78  07 January 2005 allow other text type formats to be merged.
        Print #FileFile, "By Faster than Sight"    ' 79 11Jun2010 display at the top of the video screen
        Print #FileFile, "PROMPTDETAILS"    ' 80  11Jun2010 hide the prompt details if large font slowed text display
        Print #FileFile, "noRAND_GROUP"    ' 81 27Aug2010 When in random mode and rand_group play videos randomly by group ie barrel racer etc
        Print #FileFile, "SEQUENTIAL MANUAL"    ' 82 19Aug2011 allow for the file to use control1.txt similar to cmd(45)
        Print #FileFile, "noOFFSET==-0.5000"    ' 83  02Nov2011 The offset value for video playback -1000 is one second before the start==value 01Oct2011
        Print #FileFile, "d:\bad_mp3\"       ' 84 14Jul2016 photo copy directory for mci error see cmd(23) for other default
        Print #FileFile, "003"    ' 85 17Dec2017 smlbud the video option to move 3/1000 second and much more or les
        Print #FileFile, "0300"    ' 86  17Dec2017 bigbud the video option to do fast forward with smlbud
        Print #FileFile, "011"    ' 87  17Dec2017 maxbud the number of times the smlbud plays 60 is a lot...
        Print #FileFile, ""    ' 88
        Print #FileFile, ""    ' 89
        Print #FileFile, ""    ' 90
        Print #FileFile, ""    ' 91
        Print #FileFile, ""    ' 92
        Print #FileFile, ""    ' 93
        Print #FileFile, ""    ' 94
        Print #FileFile, ""    ' 95
        Print #FileFile, ""    ' 96
        Print #FileFile, ""    ' 97
        Print #FileFile, ""    ' 98
        Print #FileFile, ""    ' 99
        Print #FileFile, ""    ' 100
        Close FileFile

Return

line_30600:         'january 09 2001
        'crop incomming file to 120 or some other line length
        xtemp = Cmd(19)
        GoSub line_16000
'            save_line = "16100-7"  '18 March 2007 testing only only
        GoSub line_16100    'january 09 2001
        crop_len = Val(Cmd(21)) - 2    'the default
        'the crop_len is reduced by 2 as the display portion of program always loads
        'a space in front. One wants to do the break 1 space before that so that
        'display doesn't do another line break  january 13 2001
        xtemp = InputBox("crop line length prompt", "test", CStr(crop_len), xx1 - offset1, yy1 - offset2) '
        crop_len = Val(xtemp)
        temp2 = 0
        Open TheFile For Input As #OutFile

        save_line = "30610"
        dblStart = Timer    'get the start time
line_30610:
        Line Input #OutFile, aaa
line_30612:
    'might want to remove other odd ball characters here too todo **vip**
    tt = InStr(aaa, Chr(9)) 'check for tabs
    If tt = 0 Then
        GoTo line_30620
    End If
    'change any tabs to 4 spaces
    aaa = Left(aaa, tt - 1) + "    " + Right(aaa, Len(aaa) - tt)
    GoTo line_30612

line_30620:
    'below get rid of all trailing spaces
    II = Len(aaa)
    If Right(data_aaa, 1) = " " Then
        aaa = Left(aaa, II - 1)
        GoTo line_30620
    End If
line_30625:
    II = Len(aaa)
    tt = 0
    If II <= crop_len + over_lap Then GoTo line_30630
    tt = InStr(crop_len, aaa, " ")
    'if a - is going to split a word make sure that there are no nearby spaces below
    If (tt = 0 Or tt > crop_len + over_lap) And Mid(aaa, crop_len - 1, 1) = " " Then tt = crop_len - 1
    If (tt = 0 Or tt > crop_len + over_lap) And Mid(aaa, crop_len - 2, 1) = " " Then tt = crop_len - 2
    If (tt = 0 Or tt > crop_len + over_lap) And Mid(aaa, crop_len - 3, 1) = " " Then tt = crop_len - 3
    If (tt = 0 Or tt > crop_len + over_lap) And Mid(aaa, crop_len - 4, 1) = " " Then tt = crop_len - 4
    If tt = 0 Then GoTo line_30628
    If tt > crop_len + over_lap Then GoTo line_30628
    Print #ExtFile, Left(aaa, tt - 1)
    aaa = Mid(aaa, tt + 1)
    GoTo line_30625
line_30628:
    Print #ExtFile, Left(aaa, crop_len); "-"
    temp2 = temp2 + 1
    aaa = Mid(aaa, crop_len + 1)
    GoTo line_30625
line_30630:
        Print #ExtFile, aaa
        GoTo line_30610
line_30640:
    dblEnd = Timer      'get the end time
    Print "     elap="; Format(dblEnd - dblStart, "#####0.000")
    Print "rename "; FileExt; " as "; TheFile
    tt1 = UCase(InputBox("total changes=" + CStr(temp2) + " rename files Y/N <N>", "Rename Prompt", , xx1 - offset1, yy1 - offset2)) '
    If tt1 <> "Y" Then
        tt1 = "N"
    End If
    If UCase(tt1) = "N" Then
        Close OutFile, ExtFile
        GoTo line_30648
    End If
    Close OutFile, ExtFile
    DoEvents
    save_line = "30640"
'    Kill "c:\oldfile.txt"
    Kill "oldfile.txt"  'december 3 2000
    DoEvents
line_30645:
'     tt1 = InputBox("minor pause", , , 4400, 4500) 'TESTING ONLY
    DoEvents
'    Name TheFile As "c:\oldfile.txt"
    Name TheFile As "oldfile.txt"   'december 3 2000
    DoEvents
    Name FileExt As TheFile
    'rename files here if Y entered
line_30648:
Return              'january 09 2001
'16Apr2014 main auto catalog routine fast forward and rewind setup
line_30700:         'april 01 2001   get files routine for copy from cd? to disc
'prompt for the various options first giving defaults
    'prompt for the source of the files maybe with sub-directories etc  Left(App.Path, 3)
'18Jun2012
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = App.Path
        GoTo line_30700a
End If                  '18Jun2012
    Test1_str = InputBox("input directory ", "data source (your input required) ", Cmd(35), xx1 - offset1, yy1 - offset2) 'april 01 2001
'    Test1_str = InputBox("input directory ", "data source (your input required) ", App.Path, xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30700a:
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    indir = Test1_str
    If Right(indir, 1) <> "\" Then indir = indir + "\"  'june 17 2001
    If diryes = "DIR" Then GoTo line_30701    'february 20 2002
    If Cmd(35) <> indir Then
        GoSub Control_28000        'october 07 2001
        Cmd(35) = indir     'june 26 2001
        GoSub line_30800    'kill and update the control file
    End If
line_30701:
    Test1_str = UCase(Test1_str)
    If Test1_str = "C:\*" Then
        indir = Left(indir, 3)      'allow for complete directory of C:\
        GoTo line_30702
    End If              'april 11 2001
    'ensure that they arn't allowed to find every *.jpg on C drive? Below....
    'allow them to do a complete directory of any device / drive june 10 2001
'june 10 2001    II = InStr(1, Test1_str, "C:\")
    II = 0      'june 10 2001
    If II <> 0 Then
        If InStr(II + 3, Test1_str, "\") = 0 Then
            Test1_str = InputBox("Wouldn't want to do all c:\", "never all of c drive", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            GoTo line_30700
        End If
    End If
line_30702:
    'when one wants just to catalogue them where they reside rather than copying them etc
    If diryes = "DIR" Then GoTo line_30704      'february 18 2002
    '18Jun2012 dougheredoug
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "Y"
        GoTo line_30702a
End If                  '18Jun2012
    Test1_str = InputBox("inplace catalogue ", "no copy (Y/N)<Y>", "Y", xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30702a:    '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    inplace = UCase(Test1_str)
    outdir = ""
    If inplace = "Y" Then GoTo line_30704
    
    'make sure the destination directory is not one of those that are listed for the *.jpg directory
    'todo **vip** see above
    Test1_str = InputBox("destination directory ", "storage area (input needed here)", Cmd(23), xx1 - offset1, yy1 - offset2) 'april 01 2001
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    outdir = Test1_str
line_30703:
'may 28 2001    Test1_str = InputBox("2 leading characters for file ", "delete to keep file name", "x-", xx1 - offset1, yy1 - offset2) 'april 01 2001
    Test1_str = InputBox(" keep original file name Y/N  <Y> (pict for pict00001 & pict00002 etc)", "Output file name prompt   ", "Y", xx1 - offset1, yy1 - offset2)
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    leaddidg = Test1_str
    delete_file = "N"       'june 30 2001
    If Test1_str = "Y" Then leaddidg = ""   'may 28 2001
'may 28 2001
'june 30 2001
    Test1_str = InputBox(" DELETE original file Y/N  <N> ", "Delete original file   ", "N", xx1 - offset1, yy1 - offset2)
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    delete_file = UCase(Test1_str)

'    If Len(leaddidg) <> 2 Then GoTo line_30703 'enforce 2 characters april 05 2001
line_30704:
'17Mar2014 ask for start point offset 0 1 2 etc
    pad_time = 0                    '17Mar2014
     If Mergem = "FF" Then         '17Mar2014
    Test1_str = InputBox("FastForward start point offset ", "0 1.5 2<0>", "0", xx1 - offset1, yy1 - offset2) '17Mar2014
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    pad_time = Val(Test1_str) * 1000  '17Mar2014
    End If          '17Mar2014
'17Mar2014
'19 September 2004    temps = "ALL, MPG, MP3, JPG, BMP, TXT, WMV, AVI"   'february 23 2002
'the merge should also allow frm cls bas files as well ie the vb code merge
'28 November 2004    temps = "MPG, MP3, JPG, BMP, WMV, AVI (x to exit)"   'february 23 2002
    temps = "MPG, MP3, JPG, BMP, WMA, WMV, AVI, GIF, ALL, (x to exit)"   'february 23 2002
'07 January 2005 allow control file to contain what formats to merge
'07 January 2005    If Mergem = "YES" Then temps = "TXT, " + temps  '19 September 2004
    If diryes = "DIR" Then temps = "ALL, " + temps  '20 September 2004
    If Mergem = "YES" Then temps = "TXT, " + Cmd(78)  '07 January 2005
'18 August 2004    If diryes = "DIR" Then temps = temps + " <A> for All"    'february 23 2002
    xtemp = Cmd(36)     'february 23 2002
    If Mergem = "YES" Then xtemp = "TXT"    '07 January 2005
    If Mergem = "FF" Then xtemp = "MPG"     '21Mar2014
'20 September 2004    If xtemp = "ALL" Then xtemp = "JPG"     '19 September 2004 (just because its out there in older versions)
'    If xtemp = "ALL" And diryes <> "DIR" Then xtemp = "mp3"     '27Jun2016 (just because its out there in older versions)
    If xtemp = "ALL" And diryes <> "DIR" Then xtemp = "mpg"     '06Mar2018 (just because its out there in older versions)
'27 may 2002    xtemp = "A"         'february 23 2002
If Left(UCase(App.EXEName), 13) = "CATMYDRIVEJPG" Then
        Test1_str = "JPG"
        GoTo line_30704a
End If                  '18Jun2012
If Left(UCase(App.EXEName), 13) = "CATMYDRIVEMP3" Then
        Test1_str = "MP3"
        GoTo line_30704a
End If                  '23Jun2012
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "MPG"
        GoTo line_30704a
End If                  '18Jun2012
    Test1_str = InputBox("file type / extension ", temps, xtemp, xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30704a:
    Test1_str = UCase(Test1_str)
    If Test1_str = "A" Then Test1_str = "ALL"   '07 January 2005

'09 July 2003 over-ride the PHOTO_DETAIL setting
    '09 July 2003 "photo_detail" must be set for the data in the TAG to be used... vip
    If Test1_str = "MP3" Then
        Cmd(56) = "PHOTO_DETAIL"            '09 July 2003 needed to get the info out for now anyway
        detailyn = "PHOTO_DETAIL"           '18 November 2004
        Cmd(49) = "NORANDOM"    'this seemed to cause some problems too...
    End If                                  '09 July 2003
    If Test1_str = "X" Or Test1_str = "E" Then GoTo End_32000
    If InStr(Cmd(78), Test1_str) <> 0 And Len(Test1_str) = 3 And Mergem = "YES" Then GoTo line_30705   '07 January 2005
    If InStr(Cmd(78), Test1_str) <> 0 And Len(Test1_str) = 2 And Mergem = "YES" Then GoTo line_30705   '19 November 2006
    If Test1_str = "TXT" And Mergem = "YES" Then GoTo line_30705   '19 September 2004
'07 January 2005 add logic to have merge of text file types to come from control file ie htm, cls, frm etc
    If Mergem = "YES" Then GoTo line_30704      '07 January 2005 do not merge non text formats ever.
    If Test1_str = "VOB" Then GoTo line_30705   '15 December 2006
    If Test1_str = "MPG" Then GoTo line_30705   '19 September 2004
    If Test1_str = "MP3" Then GoTo line_30705   '19 September 2004
    If Test1_str = "JPG" Then GoTo line_30705   '19 September 2004
    If Test1_str = "BMP" Then GoTo line_30705   '19 September 2004
    If Test1_str = "WMA" Then GoTo line_30705   '03Jul2016
    If Test1_str = "WMV" Then GoTo line_30705   '19 September 2004
    If Test1_str = "AVI" Then GoTo line_30705   '19 September 2004
    If Test1_str = "GIF" Then GoTo line_30705   '28 November 2004
    If Test1_str = "LOG" Then GoTo line_30705   '27 November 2006
    If Test1_str = "ALL" Then GoTo line_30705   '20 September 2004
     
'11Jul2016 all get to do gf     GoTo line_30704                             '19 September 2004

line_30705:             '19 September 2004 only allow the ones above
'19 September 2004 the option "ALL" for dirtype is no longer in use.
'   it is confusing to allow formats that the program can not display, therefore no more ALL

'07 January 2005    If Test1_str = "A" Then Test1_str = "ALL"   'february 23 2002
    dirtype = Cmd(36)        'february 23 2002
    If Test1_str = "ALL" Then dirtype = "ALL"     'february 23 2002
    If Test1_str = "ALL" Then
        filetype = "*"
    Else
        filetype = Test1_str
    End If 'february 23 2002
    If Cmd(36) <> dirtype Then
        GoSub Control_28000        'october 07 2001
        Cmd(36) = dirtype     'june 30 2001
        GoSub line_30800    'kill and update the control file
    End If
    If diryes = "DIR" Then
'february 23 2002        filetype = "*"      'february 18 2002
        minsize = -1
        maxsize = 99999
        maxsize = 9999999999#             '08 june 2003
        autobuild = "Y"
        savepath = "N"
        skipyesno = "N"
        inplace = "Y"
        GoTo line_30706     'february 18 2002
    End If
    'can use * for filetype above april 11 2001
    
    'should probably strip out any . periods????
    'use FileLen when checking file size
    'min size handy for skippin thumbnail picts ie those less than 20,000
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = ".5"
        GoTo line_30705a
End If                  '18Jun2012
    Test1_str = InputBox("minimum size ", "smallest file (min =0)", ".5", xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30705a:        '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    minsize = Val(Test1_str)
    'max size handy to skip extra large files or for seeing those under 20,000 too
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "999999999"
        GoTo line_30705b
End If                  '18Jun2012
    Test1_str = InputBox("maximum size ", "largest file", "999999999", xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30705b:        '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    maxsize = Val(Test1_str)
    'the auto copy is handy for when just a backup and initial file create wanted ie no intereaction
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "Y"
        GoTo line_30705c
End If                  '18Jun2012
    Test1_str = InputBox("auto build ", "selection prompts (Y/N)<Y>", "Y", xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30705c:        '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    autobuild = UCase(Test1_str)
    
    'no need to save path really if inplace catalogue unless for searching
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "Y"
        GoTo line_30705d
End If                  '18Jun2012
    Test1_str = InputBox("original path in text line ", "track original (Y/N)<Y>", "Y", xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30705d:        '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    savepath = UCase(Test1_str)
    
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        Test1_str = "N"
        GoTo line_30705e
End If                  '18Jun2012
    Test1_str = InputBox("skip duplicate file sizes ", "skip duplicates (Y/N)<N>", "N", xx1 - offset1, yy1 - offset2) 'june 09 2001
line_30705e:        '18Jun2012
    If UCase(Test1_str) = "X" Or UCase(Test1_str) = "E" Then GoTo End_32000
    skipyesno = UCase(Test1_str)
    
    'log the file locations to the search text file replace.txt etc
'    xtemp = Cmd(19)
line_30706:             'february 18 2002
    filereason = "Search link text file name"
    xtemp = Cmd(37)     'june 30 2001
    If diryes = "DIR" Then xtemp = "directory.txt"     'february 24 2002
    If Mergem = "YES" Then xtemp = "merge.txt"          'february 24 2002
    GoSub line_16000          'february 24 2002 moved up
        oopen = "WRITE"     'september 02 2001
        'february 17 2002 add xtemp name below
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        oopen = "O"
        GoTo line_30706a
End If                  '18Jun2012
        oopen = InputBox(" append(A) or write(O) <over-write> " & FileExt, "Append or New Prompt   ", "O", xx1 - offset1, yy1 - offset2)
line_30706a:        '18Jun2012
        oopen = UCase(oopen)
        If oopen <> "O" Then oopen = "APPEND"
        If oopen <> "APPEND" Then
            oopen = "WRITE"  'september 02 2001
        End If
    If diryes = "DIR" Then
'february 24 2002        xtemp = "directory.txt"
'february 24 2002       FileExt = xtemp
'february 25 2002        oopen = "WRITE"         'february 18 2002
        GoTo line_30708
    End If                      'february 18 2002
    If Cmd(37) <> FileExt Then
        GoSub Control_28000       'october 07 2001
        Cmd(37) = FileExt     'june 30 2001
        GoSub line_30800    'kill and update the control file
    End If
line_30708:                  'february 18 2002
            'following lines moved down from above the if february 18 2002
'            save_line = "16100-8"  '18 March 2007 testing only only
    GoSub line_16100    'open the replace.txt for output
'            save_line = "16100-8a"  '18 March 2007 testing only only
    oopen = "write"     'september 02 2001
    dblStart = Timer      'get the start time
'    If diryes = "DIR" Then
'        testtest = InputBox("RIGHT AFTER FILE OPEN " + FileExt, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'    End If
    filereason = ""
    ttt1 = 0
    tt = 1
    temp_cnt = 0        '08 July 2003
    Cls
    cript1(tt) = indir
line_30710:
    II = ttt1 + 1
    III = tt
'08 July 2003 see what is happening to Peter tran tap's computer when a dir done
    temp_cnt = tt     '08 July 2003 testing only can be removed when done
    Print "testing directory " + CStr(temp_cnt) + " " + CStr(Len(Trim(cript1(tt)))) + " " + cript1(tt) '08 July 2003 testing
    Print "-"       '08 July 2003 testing to clear above print line to screen meebee
        new_delay_sec = 1   '18 August 2004
'            save_line = "16100-8b"  '18 March 2007 testing only only
        GoSub line_30300    '18 August 2004
'            save_line = "16100-8c"  '18 March 2007 testing only only
'        delay_sec = 2
'        GoSub line_30000        'do a temporary delay sleep wait while testing for now handy
'25 July 2003        delay_sec = Format(Val(Cmd(27)), "###0.0000")
        delay_sec = Val(Cmd(27))
    If temp_cnt Mod 20 = 0 Then
        Cls
    End If          '08 July 2003
'           save_line = "16100-8d"  '18 March 2007 testing only only
    For JJ = II To III
            save_line = "see folder " + CStr(cript1(JJ)) '18 March 2007 testing only only
'18 March 2007 somehow my favorites are causing this problem just skip em for now
' one other place where favorites needs to be skipped
'        frmproj2.Caption = " working folder " + cript1(JJ) '01 march 2010
    If InStr(1, UCase(cript1(JJ)), "FAVORITES") <> 0 Then GoTo line_30711 '18 March 2007
    If InStr(1, UCase(cript1(JJ)), "LIVEKERNALREPORTS") <> 0 Then GoTo line_30711 '02 March 2010
           save_line = "16100-8d"  '03 March 2010 to skip all bad directories
        
    thedir = Dir(cript1(JJ), vbDirectory)
'    Test1_str = InputBox("In the get folders " + cript1(JJ) + " " + CStr(II) + " " + CStr(III) + " " + CStr(tt), "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'    If Test1_str = "x" Or Test1_str = "e" Then GoTo End_32000
    Do While thedir <> ""
'    Print "testing directory " + thedir, CStr(temp_cnt) '08 July 2003 testing
'    Print "testing directory " + CStr(temp_cnt)  '08 July 2003 testing
    
    If thedir <> "." And thedir <> ".." Then
        If (GetAttr(cript1(JJ) & thedir) And vbDirectory) = vbDirectory Then
'            pp = FileLen("a:\temp.txt")    'testing
            ' Display entry only if it
            DoEvents
'01 july 2003 activate following line
'    Test1_str = InputBox("folders testing doug" + cript1(JJ) + "*" + outdir + "*" + thedir + "*" + indir + "*" + photo_dir, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
            If cript1(JJ) + thedir + "\" <> outdir Then
                tt = tt + 1
                cript1(tt) = cript1(JJ) + thedir + "\" 'save each folder and subfolder
    
    '27 june 2003 the few following lines were used to test for directory errors on partitioned disk
'    i = i + 1
'    If i Mod 20 = 0 Then
'        Cls
'        i = 0
'    End If
'    Print "test#1="; CStr(tt); "  "; cript1(tt)         '27 june 2003
'    Print "   "
'    Print "   "         '27 june 2003
                If tt Mod 100 = 0 Then
                    Print "working folders="; tt, cript1(JJ)
                End If
            End If      'just in case the outdir is in the same folder???
                           'it would go on and on coping in the same directory
        End If  ' it represents a directory.
    End If
'18 March 2007    If tt > 9999 Then
'04 March 2010 changed from 20000 to 25000 max
    If tt > 24999 Then
        Test1_str = InputBox("Max folders of 25000 " + CStr(tt), "needs changing ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
        GoTo line_30713
    End If              'for now set the max number of folders to 400
    thedir = Dir        'get next directory command
    Loop
line_30711:             '18 March 2007 skip my favorites files from cataloging seems to be a problem
'           save_line = "16100-8d"  '03 March 2010 to skip all bad directories
    save_line = "30711" '05 March 2010
Next JJ
'            save_line = "16100-8e"  '18 March 2007 testing only only
    ddd = 0     'may 28 2001
    ttt1 = III
    save_line = "30711a" '05 March 2010
    If III <> tt Then GoTo line_30710
line_30713:
'25 june 2003 somewhere along the line here disk partitions do not seem to work???
'now start looking for the files in each of these directories / folders
    II = 0
    Cls
    For JJ = 1 To tt
    save_line = "30711b" '05 March 2010
    thedir = Dir(cript1(JJ) + "*." + filetype, vbDirectory) 'directory command
    save_line = "30711bb" + CStr(JJ) '05 March 2010
    Print "="; cript1(JJ); "="  'april 11 2001 and 08 July 2003 checks...
    If JJ Mod 20 = 0 Then
        new_delay_sec = 0.1   '18 August 2004
        GoSub line_30300    '18 August 2004
'        delay_sec = 0.2
'       GoSub line_30000                'do a temporary delay sleep wait while testing for now handy
        Cls
'25 July 2003        delay_sec = Format(Val(Cmd(27)), "###0.0000")
        delay_sec = Val(Cmd(27))
    End If          '08 July 2003
    save_line = "30711c" '05 March 2010
    Do While thedir <> ""
'    Do While thedir <> "." And thedir <> ".."
    If thedir = "." Then GoTo line_30720
    If thedir = ".." Then GoTo line_30720
    pp = FileLen(cript1(JJ) + thedir)  'testing
'02Feb2012 maybe get the file length in seconds if a mpg file here and use it in addition to pp
'add a control file element to do this video_length only when requested it does a lot of flickering etc
        elapse_start = Timer            '06Mar2018
        video_length = 0            '03Feb2012
        Test2_str = ""              '04Feb2012
        Test2_str = " bytes=" + CStr(pp)  '06Feb2012
        '06Feb2012    If filetype = "MPG" And Mergem = "VL" Then
    If (filetype = "MPG" And Mergem = "VL") Or (filetype = "MPG" And Mergem = "FF") Or (filetype = "MPG" And Mergem = "GOLF") Then
        'testtest = "filetype=" + filetype + " video_length="
        'testtest = InputBox("testing filetype " + testtest + " " + cript1(JJ) + thedir, "get files ", , xx1 - offset1, yy1 - offset2) '02Feb2012
        'frmproj2.Caption = "(Finding Video length)" + cript1(JJ)   'testing only
        Last$ = frmproj2.hWnd & " Style " & &H40000000
        Pict_file = cript1(JJ) + thedir
        lresult = getshortpathname(Pict_file, sshortfile, Len(sshortfile))  '04Feb2012
        Pict_file = Left$(sshortfile, lresult)                              '04Feb2012
    save_line = "30711d" + Pict_file + filetype + " " + Mergem '30Mar2014
        todo$ = "open " + Pict_file + " Type MPEGVideo Alias video1 wait parent " & Last$
        result1 = mciSendString(todo$, returnstring1, 1024, 0)   '12 December 2004a
        DoEvents
        i = mciSendString("put video1 window at -1 -1 " + Cmd(51) + " " + Cmd(52), 0&, 0, 0)
        DoEvents
    save_line = "30711e" + Pict_file + filetype + " " + Mergem '30Mar2014
        i = mciSendString("status video1 length wait", mssg, 255, 0)
        DoEvents
        temp3 = InStr(mssg, Chr$(0))
        video_length = Val(Left(mssg, temp3 - 1))
'        testtest = "filetype=" + filetype + " video_length=" + " result1=" + CStr(result1) + " pp=" + CStr(pp)
'        testtest = InputBox("testing filetype " + testtest + " " + cript1(JJ) + thedir + " mssg=" + mssg, "get files ", , xx1 - offset1, yy1 - offset2) '02Feb2012
'        frmproj2.Caption = "(Finding Video length)" + cript1(JJ)   'testing only
        i = mciSendString("close video1", 0&, 0, 0)
        DoEvents        '30Mar2014
        'need a close on video1
        elapse_end = Timer            '06Mar2018
    Test2_str = " len=" + CStr(video_length) + " bytes=" + CStr(pp) + " elap=" + Format(elapse_end - elapse_start, "#####0.000") '06Mar2018 add elapsed
    End If
'02Feb2012
    If pp = last_pp And skipyesno = "Y" Then
        skip_pp = skip_pp + 1
        GoTo line_30720         'ken had a bunch of duplicates
    End If                  'may 10 2001
    'need the file datecreated and modified changed dates april 11 2001
'    If diryes = "DIR" Then
'        testtest = InputBox("Testing dir step 1 " + FileExt + Format(pp, "########"), "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'        If testtest = "X" Then GoTo What_50
'    End If
'    Test1_str = InputBox("maximum size ", "largest file", "999999999 " + Format(pp, "##############") + Test2_str, xx1 - offset1, yy1 - offset2) '21Mar2014 test only
    If minsize * 1000 > pp And Mergem <> "FF" Then GoTo line_30720  'skip it too small
    If maxsize * 1000 < pp And Mergem <> "FF" Then GoTo line_30720  'skip it too big
    If minsize * 1000 > video_length And Mergem = "FF" Then GoTo line_30720  'skip it too small
    If maxsize * 1000 < video_length And Mergem = "FF" Then GoTo line_30720  'skip it too big
    tt1 = ""
    If autobuild = "Y" Then GoTo line_30715
    If img_ctrl = "YES" Then
        Set Image1.Picture = LoadPicture(cript1(JJ) + thedir)
    Else
            If Not mixx Then Set Picture = LoadPicture(cript1(JJ) + thedir)           '12Feb2017
'        Set Picture = LoadPicture(cript1(JJ) + thedir)
    End If
    tt1 = InputBox("select picture Y/N <Y> " + cript1(JJ) + thedir + " " + CStr(pp), "Or enter photo description", , xx1 - offset1, yy1 - offset2)
    If UCase(tt1) = "N" Then GoTo line_30720   'they didn't want it
    If UCase(tt1) = "E" Or UCase(tt1) = "X" Then GoTo End_32000
line_30715:
'    If diryes = "DIR" Then
'        testtest = InputBox("Testing dir step 2 " + FileExt, "get files ", , xx1 - offset1, yy1 - offset2) 'april 01 2001
'    End If
'18 March 2007 need to skip anything from the favorites file vip
    If InStr(1, UCase(cript1(JJ)), "FAVORITES") <> 0 Then GoTo line_30730 '18 March 2007
    If save_line = "16100-8d" Then GoTo line_30730  '05 March 2010
    Test1_str = cript1(JJ) + thedir     'save the original path for search and reference
    If savepath <> "Y" Then Test1_str = " "
    newname = Test1_str
'may 28 2001    If inplace <> "Y" Then newname = outdir + leaddidg + CStr(JJ) + thedir
    If inplace <> "Y" Then
        ddd = ddd + 1
        newname = outdir + leaddidg + Format(ddd, "00000") + "." + filetype 'may 28 2001
    End If      'may 28 2001
    If inplace <> "Y" And leaddidg = "" Then newname = outdir + thedir  'april 05 2001
    If savepath <> "Y" Then Test1_str = ""
'right in here prompt for text description of the picture
    indates = ""
    If (GetAttr(cript1(JJ) & thedir) And vbDirectory) = vbDirectory Then GoTo line_30718
'   If thedir = "COMMAND.COM" Then GoTo line_30718
'   If thedir = "SCANDISK.LOG" Then GoTo line_30718
'03 November 2004    If dirdates = "Y" Then
    If dirdates <> "Yxyxx" Then             '03 November 2004 I want it on the gf for "d" display
       Set fs = CreateObject("Scripting.FileSystemObject")
       Set f1 = fs.GetFile(cript1(JJ) + thedir)
'if the date created below is unknown it fails ie not using it
'       indates = " create=" & f1.DateCreated & " mod=" & f1.DateLastModified & " acc=" & f1.DateLastAccessed
'23Jul2010       indates = " mod=" & f1.DateLastModified & " " + CStr(pp)
       indates = " "        '23Jul2010
'        Print "testing date="; indates
    End If
'    Print "testing date="; f1.DateCreated, f1.DateLastModified, f1.DateLastAccessed
line_30718:
    
'24 june 2003
    If InStr(1, UCase(Test1_str), ".MP3") <> 0 Then
        FileFile = FreeFile
        lresult = getshortpathname(Test1_str, sshortfile, Len(sshortfile))
        temptemp = Left$(sshortfile, lresult)
        Open temptemp For Binary Access Read As FileFile    '08 july 2003
'        Open temptemp For Binary As FileFile       '08 july 2003 just in case protection is a problem
'        Print "print one " + Test1_str     '08 July 2003 test
        
        Get FileFile, LOF(1) - 127, temptag

'        Print "print onea " + temptag.header + temptag.songtitle   '08 July 2003 test
        
        If UCase(temptag.header) = "TAG" Then
            temptemp = Trim(temptag.songtitle) + Trim(temptag.artist) _
                + Trim(temptag.album) + Trim(temptag.year)
            temptemp = Trim(temptemp) + Trim(temptag.comments) '+ " genre=" + temptag.genre
            If Len(Trim(temptemp)) > 12 Then
                Test1_str = Test1_str + " * " + Trim(temptemp)
                indates = ""
            End If    '25 june 2003 if there is not much in the header use directory info instead
'        Print temptag.header
'        Print "title " + temptag.songtitle  '08 July 2003
'         Print Test1_str
        End If
        Close FileFile
'        Print "print two " + Test1_str     '08 July 2003 test
        
    End If      '24 june 2003
    
'june 10 2001
'    If UCase(indir) <> "C:\" And dirdates <> "Y" Then
    If dirdates <> "Y" Then
'03 November 2004            Print #ExtFile, "photo " + Test1_str + " " + tt1
 '15Sep2010 put out the "BREAK" record here if size is small
    If pp < 1500000 Then
        If seq > 0 Or grp = 0 Then      '21Sep2010 this code is for 2 break records in a row only 1 grp bump
'27Jun2016 re the break only if
'04Jul2016        If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK"     '27Jun2016
        grp = grp + 1
        seq = 0
        End If                          '21Sep2010 tested and it works (copying videos in mass changes the seq??) 1 at a time worked??
    End If              '15Sep2010
    If pp > 1500000 Then            '15Sep2010
        seq = seq + 1
    End If
        If grp > 0 Then grpstr = "GRP#" + CStr(grp) + " SEQ#" + CStr(seq) + " "      '15Sep2010
'15Sep2010            Print #ExtFile, "photo " + Test1_str + " " + tt1 + " " + indates
'29Jan2012            Print #ExtFile, "photo " + grpstr + Test1_str + " " + tt1 + " " + indates
'29Jan2012 added the len= value so that another program can use this value to set up a fast forward file if ever
    '06Feb2012 put out multiple lines for fast forward option here below is the first line of the matched pair
'21Mar2014    If video_length <> 0 And 5 * video_length > pp Then
    If video_length <> 0 And 5 * video_length > maxsize Then
        Test2_str = Test2_str + " **video length warning**" '07Feb2012
        GoTo no_ff                  '07Feb2012
    End If                                      '07Feb2012
    If Mergem <> "FF" Then GoTo no_ff       '06Feb2012
        nocount = 1
'04Jul2016         If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK"     '27Jun2016
'        Print #ExtFile, "BREAK"        '06Feb2012
'17Mar2014 put a dummy set of lines to play the complete video if desired
'17Mar2014 remove bytes from all prints       Print #ExtFile, "phooto dummy" + " " + grpstr + " " + tt1 + " " + indates + Test2_str '04Feb2012
'        Print #ExtFile, "phooto dummy" + " " + grpstr + " " + tt1 + " " + indates + Test2_str  '04Feb2012
'        Print #ExtFile, "yyy." + Test1_str; indates '06Feb2012
'09Feb2012        Print #ExtFile, "photo pause wait=" + CStr(fast_forward) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '04Feb2012
'put the first segment play length to 1 second. It seemed to have a problem with the big screen and my laptop when too short (it failed)
'17Mar2014 play the first 5 seconds as an intro with date time and maybe explanation with no pause
'        Print #ExtFile, "photo wait=5" + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '04Feb2012
'        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
more_ff:
'17Mar2014        lresult = nocount * hold_sec * 1000            'use the tt value for skip length / start point
        lresult = (nocount - 1) * hold_sec * 1000 + pad_time      'use the tt value for skip length / start point
        If video_length * 100 > pp And pp > 1 Then
            Print #ExtFile, "warning invalid video length " + grpstr '27Mar2014
            GoTo no_ff
        End If      '27Mar2014
        If lresult >= video_length - 100 Then GoTo last_ff
    If rewind = "YES" Then          '12Apr2014
      If nocount = 1 Then
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
        Print #ExtFile, "photo RW continue wait=" + CStr(fast_forward) + " start==" + CStr(lresult) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '06Feb2012
      Else
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
        Print #ExtFile, "photo RW continue pause wait=" + CStr(fast_forward) + " start==" + CStr(lresult) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '06Feb2012
      End If          '14Apr2104
    End If                          '12Apr2014
    If rewind <> "YES" Then         '12Apr2014
    If nocount = 1 Then          '17Mar2014 on first segment do not use continue
        Print #ExtFile, "photo FF pause wait=" + CStr(fast_forward) + " start==" + CStr(lresult) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '06Feb2012
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
    Else
        Print #ExtFile, "photo FF continue pause wait=" + CStr(fast_forward) + " start==" + CStr(lresult) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '06Feb2012
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
    End If      '17Mar2014
    End If      '12Apr2014
    
        nocount = nocount + 1
        GoTo more_ff            '06Feb2012
last_ff:
        If rewind <> "YES" Then         '12Apr2014
        Print #ExtFile, "photo FF continue start==" + CStr(video_length - (fast_forward * 1000)) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '05Feb2012
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
'04Jul2016         If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK"     '27Jun2016
'27Jun2016        Print #ExtFile, "BREAK"        '06Feb2012
        Else
        Print #ExtFile, "xxx." + Test1_str; indates '06Feb2012
        Print #ExtFile, "photo RW pause start==" + CStr(video_length - (fast_forward * 1000)) + " " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str '05Feb2012
'04Jul2016         If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK"     '27Jun2016
'27Jun2016        Print #ExtFile, "BREAK"        '06Feb2012
'16Apr2014d        ttt = "REVERSE" '16Apr2014  now reverse the file order and done
'16Apr2014d        Close #ExtFile  '16Apr2014a
'16Apr2014d        GoTo line_70    '16Apr2014
        
        End If      '12Apr2014
        GoTo ff_skip        '06Feb2012
no_ff:                  '06Feb2012
'25Feb2012 first line of the gf and golf catalog below change here re again begin and a negative start
'change this a bit to make sure that the start point ie is not less than 0 and set the begin point to twice the time of the cmd(27)
    If Mergem = "GOLF" Then
            slomo_point = video_length - (Val(Cmd(27)) * 1000) - 333   '27Feb2012 take the last 1/3 sec off
            begin_locat = video_length - (Val(Cmd(27)) * 2000) - 333
            If slomo_point < 1 Then slomo_point = 10
            If begin_locat < 1 Then begin_locat = 10
            If slomo_point < begin_locat Then slomo_point = begin_locat '01May2012
'05Jul2016 the 2 breaks here and further down element 81 must be rand_group before BREAK put in catalog file
            If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK "  'everything defaults to rand by group it is easy to deactivate by mass edit of break
            Print #ExtFile, "photo 1 pause " + grpstr + " wait=" + CStr((video_length - 200) / 1000) + " start==" + CStr(100) + " " + Test1_str + " " + tt1 + " " + indates '25Jun2010
            Print #ExtFile, "xxx." + Test1_str; indates 'this is the 2nd line of first display pairs tested #1 is above
            Print #ExtFile, "photo 2 continue pause " + grpstr + " wait=" + CStr(Val(Cmd(27))) + " begin==" + CStr(begin_locat) + " start==" + CStr(slomo_point) + " speed=125 " + Test1_str + " " + tt1 + " " + indates  '25Jun2010
            Print #ExtFile, "xxx." + Test1_str; indates 'this is the 2nd line of first display pairs tested #1 is above
            Print #ExtFile, "photo 3 continue again " + grpstr + " wait=" + CStr(Val(Cmd(27))) + " start==" + CStr(slomo_point) + " " + Test1_str + " " + tt1 + " " + indates  '25Jun2010
'2          Print #ExtFile, "xxx." + Test1_str; indates   '25Jun2010
  '05Feb2012            Print #ExtFile, "photo pause begin again " + grpstr + " start==-3000 speed=125 " + Test1_str + " " + tt1 + " " + indates + " bytes=" + CStr(pp) '25Jun2010
            Else        '25Feb2012
            If Left(Cmd(81), 10) = "RAND_GROUP" Then Print #ExtFile, "BREAK "
            Print #ExtFile, "photo " + grpstr + Test1_str + " " + tt1 + " " + indates + Test2_str    '04Feb2012
    End If              '25Jun2010
            indates = ""            '03 November 2004
    End If
    If inplace = "Y" Then Test1_str = cript1(JJ) + thedir
    If inplace <> "Y" Then Test1_str = newname
        If diryes = "DIR" Then      'february 18 2002
            If Mergem = "YES" Then indates = " append start" 'February 24 2002
            'need to replace indates above so that later diff program will work on a program level
            'need to skip the copy logic etc here if the file name same as output name  vip
'            testing = "          " + Format(pp, "000000000#") '08Sep2016
            testing = Trim(Format(pp, "          #")) '08Sep2016
            testing = Space(11 - Len(testing)) + testing
'            testing = Right(testing, Len(testing) - 9)
'            testing = Right(testing, Len(testing) - 1) '08Sep2016
            Print #ExtFile, "size="; testing; " "; Test1_str; indates '07Sep2016 add the file size to output
'07 January 2005 need to check that it is one of those allowed in Cmd(78) before merging
            If Mergem = "YES" Then GoSub line_17200 'february 24 2002
            ' need to open need to open and read the file here thru a call
            'then put and append end along with the file name at the end.
            'should just need to use one file # for the opens and closes...
        Else
            Print #ExtFile, "xxx." + Test1_str; indates 'this is the 2nd line of first display pairs tested #1 is above
        End If                      'february 18 2002 skip the xxx. on print only
    If Mergem = "GOLF" Then
'15Sep2010            Print #ExtFile, "photo start==-3000 speed=125 " + Test1_str + " " + tt1 + " " + indates   '25Jun2010
'06Feb2012            Print #ExtFile, "photo " + grpstr + "start==-3000 speed=125 " + Test1_str + " " + tt1 + " " + indates '25Jun2010
'25Feb2012            Print #ExtFile, "photo " + grpstr + "start==-3000 speed=125 " + Test1_str + " " + tt1 + " " + indates + " bytes=" + CStr(pp) '25Jun2010
  '11Mar2012          Print #ExtFile, "photo continue begin==10 " + grpstr + Test1_str + " " + tt1 + " " + indates + " bytes=" + CStr(pp)  '25Jun2010
'29Jan2012            Print #ExtFile, "xxx." + Test1_str; indates    '25Jun2010
'06Feb2012            Print #ExtFile, "xxx." + Test1_str; indates + " len=" + CStr(pp)   '25Jun2010  ***this looks wrong***
  '11Mar2012          Print #ExtFile, "xxx." + Test1_str; indates   '25Jun2010
    End If              '25Jun2010
 '04Jul2016           Print #ExtFile, "BREAK "
            '06Feb2012 come to here after the catalog for fast forward is done
ff_skip:        '06Feb2012 when done ff go to here
'11 december 2002        II = II + 1             'count of files selected
    EE = EE + 1         '11 december 2002
    If EE Mod 100 = 0 Then
        Print "working pics="; EE; " "; Test1_str; indates; " "; CStr(pp);
    dblEnd = Timer      'get the end time
    Print "  elap="; Format(dblEnd - dblStart, "#####0.000")
        'may 10 2001
    End If
'11 december 2002    If II Mod 2000 = 0 Then
    If EE Mod 2000 = 0 Then
'    frmproj2.Caption = program_info + " (cls #16)" '18Dec2013
        Cls           'clear after every 20 lines
    End If      'april 11 2001
    If inplace <> "Y" Then
                DoEvents
'25Aug2016
         If Len(photo_dir) > 1 Then FileCopy cript1(JJ) + thedir, newname               'copy the file
'        xtemp = InputBox(" copy picture to " + temps + " <Y>", "Copy Prompt   ", disp_file, xx1 - offset1, yy1 - offset2) '12Jul2016 testing
   frmproj2.Caption = program_info + disp_file '12Jul2016
 '       delay_sec = 1      'may 10 2001
 '       GoSub line_30000
                DoEvents
'Print "testing doug delete_file="; cript1(JJ) + thedir
                 If delete_file = "Y" Then
                    Kill cript1(JJ) + thedir
                    DoEvents        'june 30 2001
                 End If
    End If                      'copy file here
line_30720:
    last_pp = pp
    DoEvents
    thedir = Dir        'get next directory
    Loop
line_30730:
    save_line = "30730" '05 March 2010
    Next JJ
'12 december 2002    Test1_str = InputBox("total folders= " + CStr(III) + " total files= " + CStr(II) + " skip=" + CStr(skip_pp), "session counts " & FileExt, , xx1 - offset1, yy1 - offset2) 'april 01 2001
'18Jun2012
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        GoTo line_30730a
End If                  '18Jun2012
    Test1_str = InputBox(" total folders= " + CStr(III) + " total files= " + CStr(EE) + " skip=" + CStr(skip_pp), "session counts " & FileExt, , xx1 - offset1, yy1 - offset2) 'april 01 2001
line_30730a:
'22 june 2003 put 10 blank lines at end of file
    For JJ = 1 To 10
    Print #ExtFile, ""
    Next JJ     '22 june 2003
If Left(UCase(App.EXEName), 10) = "CATMYDRIVE" Then
        GoTo End_32000
End If                  '18Jun2012

    Close #ExtFile

Return                  'april 01 2001
'auto catalog routines above  subroutine   16Apr2014

line_30800:                     'update control.txt with changes write control file
        DoEvents
'        Kill "C:\control.txt"
    If auto_exe = "YES" Then GoTo line_30850    '04 September 2004
    If InStr(Cmd(40), Left(App.Path, 3)) = 0 And Trim(Cmd(40)) <> "" Then
  '23Jun2012a          xtemp = InputBox("app.path" + App.Path + " not found or updated " + Format(delay_sec, "###0.0000") + Cmd(40), , , 4400, 4500)  'TESTING ONLY
        Print " no control file update cmd(40)"
        GoTo line_30850
    End If              '20 july 2002 must be a writeable device in cmd(40)
'        Kill "control.txt"      'december 3 2000
        Kill control_file      '23 september 2002
        FileFile = FreeFile
        II = DoEvents
'        Open "C:\control.txt" For Output Access Write As #FileFile
'        Open "control.txt" For Output Access Write As #FileFile 'december 3 2000
        Open control_file For Output Access Write As #FileFile '23 september 2002
        For f = 1 To 100
            If Len(Cmd(f)) < 1 Then
                Cmd(f) = Space(50)
            End If
            Print #FileFile, Cmd(f)
        Next f
        Close FileFile
        GoSub Control_28000     'october 07 2001
line_30850:     '20 july 2002
'        ttt = ""               ''october 07 2001
        Return

line_30900:                 'the "HELP" routine here october 14 2001
    FileFile = FreeFile
    Open "help.txt" For Output As FileFile
    Print #FileFile, " File Prompt #1"
    Print #FileFile, "      Enter a file to use '?.txt' or select one"
    Print #FileFile, "      of the previously selected files 1-20"
    Print #FileFile, " Option Prompt #2"
    Print #FileFile, "      Display pictures, Enter text, Search text"
    Print #FileFile, "      Help, Append text, Screen Saver, Set program"
    Print #FileFile, "      control options etc etc etc"
    Print #FileFile, ""
    Print #FileFile, "        ***** Prompt #2 Options *****"
    Print #FileFile, " C context search / most common search option"
    Print #FileFile, " CC context search / more speed combination of 'S' and 'C' search"
    Print #FileFile, " E Enter more notes to the end of the current text file"
    Print #FileFile, " F Flash search display option, all lines displayed, stopping on matches"
    Print #FileFile, " Q Single line search matching case sensitive (see 'S')"
    Print #FileFile, " S single line match display Lower & Uppercase sensitive"
    Print #FileFile, "   Both Q And S are now the same If search is numeric then no uppercase shift etc"
    Print #FileFile, " Y Display the last screen of data ie Yesterday or Last data"
    Print #FileFile, " CH change to text file (uses wordpad for edits)"
    Print #FileFile, " M Minimize screen (allows start of other desktop processes)"
    Print #FileFile, " X Exit to file name prompt"
    Print #FileFile, " P1 Photo display forces (STRETCH) to fit to screen"
    Print #FileFile, " P2 Photo display without a fit to screen option (NORMAL)"
    Print #FileFile, " P  Photo display at the default display (stretch or normal)"
    Print #FileFile, " TT Screen saver timer delay ie 'tt0' 'tt5' 'tt10' 'tt.33"
    Print #FileFile, " SS screen saver display ie SSdog river flower"
    Print #FileFile, " WW screen saver display with search criteria ie flood/oldie/ice "
    Print #FileFile, " MERGE to merge.txt so all files can be searched at once ie frm"
    Print #FileFile, " DIR files directory with date and size to directory.txt"
    Print #FileFile, " DIRDATES use with GF below to list date and size of selected files"
    Print #FileFile, " GF Get File, used to capture un-catalogued pictures (use NORAND for MP3 VIP"
    Print #FileFile, " CP Copy picture option, searches catalogued pictures for copy / export to another device or folder"
    Print #FileFile, " Z Paste / Append contents of clipboard text to end of current file"
    Print #FileFile, " HELP display this help.txt info"
    Print #FileFile, "  ---------------       some lesser used options below "
    Print #FileFile, " RRR search and replace option for masive text changes"
    Print #FileFile, " RAND / NORAND Random display of picture files in SS WW P1 or P2"
    Print #FileFile, " THUMbnail for preview /sample of video or audio using time setting tt10 tt15 etc"
    Print #FileFile, " ELAP elapsed time to display after each mp3 or mpg file "
    Print #FileFile, " video / novideo turn on and off the option for video display (avi decompressor) must be installed" '10 february 2003
    Print #FileFile, " PAUSE / PA to allow for continual text display pausing for what is in the TT delay"
    Print #FileFile, " NS No Show option ie 'nsx-rated my-ex jerk'"
    Print #FileFile, " CCC change to display options in option1.txt for larger font, color changes etc"
    Print #FileFile, " XXX option to extract all displayed data to an alternate text file"
    Print #FileFile, " SKIP Don't display ie skip lines when match found ie 'skipmy-email'"
    Print #FileFile, " CRIPT Encript current text file using crip.txt as control file"
    Print #FileFile, " DECRIPT Does reversal of the cript option"
    Print #FileFile, " MYSTUFF Allows for on-line de-encription of an encripted text file"
    Print #FileFile, " LL Line length, sets line length before wrap ie 'LL100' 'LL' sets back to original length"
    Print #FileFile, " SHOWPOS Displays the line length at the end of each line"
    Print #FileFile, " SHOWASC Displays ascii value of last 10 characters on line"
    Print #FileFile, " EMAIL Reads e-mail formatted files and skips lots of control info"
    Print #FileFile, " IMPORT Takes e-mails from netscape and outlook files"
    Print #FileFile, " T Same as 'S' except that data also put to an extract file"
    Print #FileFile, " D Same as 'C' except that data also put to an extract file"
    Print #FileFile, " R Same as 'Q' except that data also put to an extract file"
    Print #FileFile, " CROP Creates extract file with line wraps ie 'CROP100'"
    Print #FileFile, " QQQ Similar to RRR used to line-feed carriage-returns from line matches"
    Print #FileFile, " VV Load clipboard with text following 'VVrepetitive data'"
    Print #FileFile, " VVV Loads secondary clipboard with data 'VVVinsert-this'"
    Print #FileFile, " GGG yet another clip area load"
    Print #FileFile, " SC Screen Capture delay (delays prompt box by 5 seconds) so alt/print-screen sequence can capture data"
    Print #FileFile, " PC Picture count  (used with 'CP' to assign picture sequence number ie 'PC20' results in pict20.jpg "
    Print #FileFile, " DISC displays username and computer name & OS version"
    Print #FileFile, " HH Sets the hilite_this element (very handy to hi-lite other info CMD(31))"
    Print #FileFile, " "        '06 December 2004 see if a blank line fixes the help display
    Print #FileFile, " HL Changes the Append_start element for display"
    Print #FileFile, " HLL Changes the Append_end element for hi-liting"
    Print #FileFile, " DETAIL NODETAIL changes cmd(56) "            '18 November 2004
    Print #FileFile, ""
    Print #FileFile, "        ***** Search Prompt #3 Options *****"
    Print #FileFile, "V or H will load clipboard data just like Ctrl V"                 '17 December 2004
    Print #FileFile, "D for todays date to be the search"
    Print #FileFile, "DX for yesterday date for search but no 01 workie"
    Print #FileFile, "otherwise the text search strings seperated by a =" + Cmd(6)       '06 December 2004
    Print #FileFile, ""
    Print #FileFile, "       ***** End of Screen #4 Prompt Options  *****"
    Print #FileFile, " P Previous for previous picture or previous match if text search"
    Print #FileFile, " PP Previous Picture passed to Browser Netscape or Explorer"
    Print #FileFile, "     allows for a picture print etc  "
    Print #FileFile, " B Back / previous page when in text search display"
    Print #FileFile, " V Same as Back moves one page before current match"
    Print #FileFile, " "
    Print #FileFile, "       ***** Detail line options and controls *****"
    Print #FileFile, " WAIT=    Over-ride wait time on a line by line basis"
    Print #FileFile, " THUMB==  Partial play using time in display seconds"
    Print #FileFile, " DOSHOW   Over-ride any noshow"
    Print #FileFile, " START==  Set the start point for a Movie or Music file"
    Print #FileFile, " FIT==    Set the photo display to fit-to-screen"
    Print #FileFile, " REG==    Display the photo without fit-to-screen"
    Print #FileFile, " "
    Print #FileFile, "       ***** Command File Control Elements *****"
    Print #FileFile, " CMD(1) Max_cnt lines per screen  ==" + Cmd(1)
    Print #FileFile, " CMD(2) Font.size text point size  ==" + Cmd(2)
    Print #FileFile, " CMD(3) Backcolor background color '3 for aqua'  ==" + Cmd(3)
    Print #FileFile, " CMD(4) Def_fore text color '15 for white' '12 for red' ==" + Cmd(4)
    Print #FileFile, " CMD(5) forecolor hi-lite color '14 for yellow' ==" + Cmd(5)
    Print #FileFile, " CMD(6) seperator '/' forward slash default search seperator ==" + Cmd(6)
    Print #FileFile, " CMD(7) dateno or dateyes for a date on every input line ==" + Cmd(7)
    Print #FileFile, " CMD(8) browser default Explorer / Netscape path ==" + Cmd(8)
    Print #FileFile, " CMD(9) paint program 'mspaint' file path ==" + Cmd(9)
    Print #FileFile, "CMD(10) text editor 'wordpad' is the default 'CH' option ==" + Cmd(10)
    Print #FileFile, "CMD(11) files.txt list of opened files in accessed order ==" + Cmd(11)
    Print #FileFile, "CMD(12) search.txt list of recent search strings in order at search prompt ==" + Cmd(12)
    Print #FileFile, "CMD(13) mspub.exe executable program ?? ==" + Cmd(13)
    Print #FileFile, "CMD(14) enterdateyes date on start of each text entry group 'E' option ==" + Cmd(14)
    Print #FileFile, "CMD(15) enddateyes date on end of each text entry group 'E' option ==" + Cmd(15)
    Print #FileFile, "CMD(16) optional .exe ??? ==" + Cmd(16)
    Print #FileFile, "CMD(17) xoffset for prompt box '11000' photo display semi-hide prompt box ==" + Cmd(17)
    Print #FileFile, "CMD(18) yoffset for prompt box '8500' photo display semi-hide prompt box ==" + Cmd(18)
    Print #FileFile, "CMD(19) replace.txt temporary file for change routines see '37'??? ==" + Cmd(19)
    Print #FileFile, "CMD(20) C/p1/email ??? ==" + Cmd(20)
    Print #FileFile, "CMD(21) line_len characters per line '82' default ==" + Cmd(21)
    Print #FileFile, "CMD(22) context_lines '6' lines displayed before a 'C' context match ==" + Cmd(22)
    Print #FileFile, "CMD(23) copy picture 'CP' default folder 'destination folder for copies' ==" + Cmd(23)
    Print #FileFile, "CMD(24) ppaste allows for vvv to paste string 'xxx.c:\family1\scn000' ==" + Cmd(24)
    Print #FileFile, "CMD(25) gpaste allows for ggg to paste string 'photo family gathering' etc ==" + Cmd(25)
    Print #FileFile, "CMD(26) 'photo' ss default ss_search element ==" + Cmd(26)
    Print #FileFile, "CMD(27) delay_sec for screen saver 'SS' option '4' seconds default ==" + Cmd(27)
    Print #FileFile, "CMD(28) noshow element ie 'creep' 'nfg' 'my-x' etc for photo display skip photos ==" + Cmd(28)
    Print #FileFile, "CMD(29) altcolor '10' default green, see element 31 which is displayed with this color ==" + Cmd(29)
    Print #FileFile, "   handy to display critical elements ie xxx. is in 31 for demo how to use this "
    Print #FileFile, "CMD(30) 'Y' unload project option required for Windows ME Y/N option ==" + Cmd(30)
    Print #FileFile, "CMD(31) hilite_this element 'XXX.' if this element exists hilite it with code in element 29 ==" + Cmd(31)
    Print #FileFile, "CMD(32) crip.txt encription file to use when 'mystuff' code  ==" + Cmd(32)
    Print #FileFile, "CMD(33) D search default at search prompt entry 'D' for today etc ==" + Cmd(33)
    Print #FileFile, "CMD(34) 5 context window 'before' see element 22 ???? ==" + Cmd(34)
    Print #FileFile, "CMD(35) indir input directory for 'GF' get file load picture option ==" + Cmd(35)
    Print #FileFile, "CMD(36) file extension 'JPG' for 'GF' get file load picture type of picture ==" + Cmd(36)
    Print #FileFile, "CMD(37) temp.txt 'see option 19' ==" + Cmd(37)
    Print #FileFile, "CMD(38) stretch_img 'stretch/normal' 'yes' or 'no' P1 or P2 photo display option ==" + Cmd(38)
    Print #FileFile, "CMD(39) autoredraw allows for screen refresh between sessions (slows down F flash display and others)==" + Cmd(39)
    Print #FileFile, "CMD(40) Allowable write drives ie A:\ C:\ ==" + Cmd(40)
    Print #FileFile, "CMD(41) Option Prompt #2 if photo file default display P1 P2 WW ==" + Cmd(41)
    Print #FileFile, "CMD(42) Option Prompt #2 if text file default display C S Q ==" + Cmd(42)
    Print #FileFile, "CMD(43) browser default Explorer / Netscape path for NT==" + Cmd(43)
    Print #FileFile, "CMD(44) text editor 'wordpad' is the default 'CH' for NT option ==" + Cmd(44)
    Print #FileFile, "CMD(45) auto run program name ie SPECIAL ==" + Cmd(45)
    Print #FileFile, "CMD(46) auto run prompt #1 (file select) ==" + Cmd(46)
    Print #FileFile, "CMD(47) auto run prompt #2 (search options) ==" + Cmd(47)
    Print #FileFile, "CMD(48) auto run prompt #3 (search string)==" + Cmd(48)
    Print #FileFile, "CMD(49) RANDOM NORANDOM on photo search ==" + Cmd(49)     '09 december 2002
    Print #FileFile, "CMD(50) Show Merged file names In Caption Area (SHOWFILES) ==" + Cmd(50) '24 december 2002
    Print #FileFile, "CMD(51) avi file size x min (1050)==" + Cmd(51) '01 february 2003
    Print #FileFile, "CMD(52) avi file size y min (775)==" + Cmd(52) '01 february 2003
    Print #FileFile, "CMD(53) video / novideo if avi installed (NOSHOWVIDEO)==" + Cmd(53) '10 february 2003
    Print #FileFile, "CMD(54) If recent files match show the duplicates or not (SHOWDUPS)==" + Cmd(54) '28 february 2003
    Print #FileFile, "CMD(55) number of lines in context file to clear (see Cmd(22)) 00 or ==" + Cmd(55)  '18 March 2003 ver=1.01
    Print #FileFile, "CMD(56) display photo detail on pictures (set when GF and MP3) (PHOTO_DETAIL)==" + Cmd(56) '22 March 2003 ver=1.02
    Print #FileFile, "CMD(57) time display on / off switch (SHOWTIME)==" + Cmd(57) '26 March 2003 ver=1.03
    Print #FileFile, "CMD(58) DEFAULT_TO_CD when this set the drive ie c:\ will default to where it is running ie E:\ F:\==" + Cmd(58) '30 May 2003 ver=1.06
    Print #FileFile, "CMD(59) Delay Timing change (should be 0.00) for mpeg movies==" + Cmd(59) '29 September 2003
    Print #FileFile, "CMD(60) Freeze delay time for mpeg movies & Photo_detail info==" + Cmd(60) '01 October 2003
    Print #FileFile, "CMD(61) Allow for set speed of playback==" + Cmd(61) '30 October 2003
    Print #FileFile, "CMD(62) Slow Motion segment duration (ie .01667)==" + Cmd(62) '08 January 2004
    Print #FileFile, "CMD(63) Show Elapsed time of video switch (slow motion only) (ie SHOWELAP)==" + Cmd(63) '11 February 2004
    Print #FileFile, "CMD(64) Pad_time (thousandsands added when in slow motion)==" + Cmd(64) '21 February 2004
                            'the above command put in because some files failed when using the get position
                            'command ie notably the niki and sparrow bird bath some weird bug??
                            'it works in most cases, and is really only needed when cataloguing etc.
    Print #FileFile, "CMD(65) Random begin for video play RANDBEG etc ==" + Cmd(65)     '23 March 2004
    Print #FileFile, "CMD(66) Interrupt video replay speed (ie 250) ==" + Cmd(66)     '10 April 2004
    Print #FileFile, "CMD(67) Thumbnail on or off ie noTHUMB or THUMB ==" + Cmd(67)     '14 April 2004
    Print #FileFile, "CMD(68) Continue at end of file (WW) ie noEOF_STOP ==" + Cmd(68)       '12 September 2004 use with background job allows them to stop
    Print #FileFile, "CMD(69) Continue after match found (P) ie noHIT_STOP ==" + Cmd(69)    '17 September 2004 use with background job allows them to stop after match found.
    Print #FileFile, "CMD(70) if FOREGROUND then background job given focus & hi-lited ==" + Cmd(70)   '17 September 2004
    Print #FileFile, "CMD(71) Allow for input from batch file ie noBATCHFILE ==" + Cmd(71)   '27 October 2004
    Print #FileFile, "CMD(72) Allow for video pictures music play list noRESULTS.TXT ==" + Cmd(72) '08 November 2004 keep a file for a play list
    Print #FileFile, "CMD(73) Switch input file when Control.txt file changes see Cmd(46) noFILESWITCH ==" + Cmd(73)    '22 November 2004
    Print #FileFile, "CMD(74) End of file command file switch (re text searches mainly) noEOFCMD ==" + Cmd(74)           '26 November 2004
    Print #FileFile, "CMD(75) Do not allow for video interrupt on cd and dvd copies VIDEOSTOP ==" + Cmd(75)   '12 December 2004
    Print #FileFile, "CMD(76) Line pause to create scrolling text noLINEPAUSE==1.5 ==" + Cmd(76) '24 December 2004
    Print #FileFile, "CMD(77) Character pause on print text display noCHARACTERPAUSE==.00333 =="; Cmd(77) '03 January 2005
    Print #FileFile, "CMD(78) Text file formats to allow during merge command TXT FRM HTM ==" + Cmd(78) '07 January 2005
    Print #FileFile, "CMD(79) the owner of this version of software ie by Faster than Sight etc==" + Cmd(79) ' 01Oct2011
    Print #FileFile, "CMD(80) PROMPTDETAILS for background large font stuff the details are annoying==" + Cmd(80) '01Oct2011
    Print #FileFile, "CMD(81) noRAND_GROUP for golf when videos are grouped by foursome this needs to be turned on==" + Cmd(81) '01Oct2011
    Print #FileFile, "CMD(82) the programs named here use control1 instead of control see cmd(45) as well==" + Cmd(82) '01Oct2011
    Print #FileFile, "CMD(83) the value here is for offset from the Start== value -1000 is 1 second off that value==" + Cmd(83) '01Oct2011
    Print #FileFile, "CMD(84) this is the default folder for MCI mp3 (see Cmd(23)) problem file==" + Cmd(84) '14Jul2016
    Print #FileFile, "CMD(85) smlbud for offset when doing special video play back ie 0030 is 30/1000 sec==" + Cmd(85) '17Dec2017
    Print #FileFile, "CMD(86) bigbud for offset when doing budge and fast forward video play back ie 0030 is 30/1000 sec==" + Cmd(86) '17Dec2017
    Print #FileFile, "CMD(87) maxbud is the number of times a budge will move ahead==" + Cmd(87) '17Dec2017
    Print #FileFile, "new feature 09 june 2002 allow for 6 search strings up from 3 (hi-lites only 3)"
    Print #FileFile, "new feature 23 june 2002 allow for hi-lites of all 6 search strings"
    Print #FileFile, "new feature 20 july 2002 allow use of non writable cd demo (do not update files)"
    Print #FileFile, "new feature 27 july 2002 at photo continue prompt allow for . to display hidden prompt on screen"
    Print #FileFile, "new feature 10 august 2002 allow for prompt #2 to default "
    Print #FileFile, "new feature 12 august 2002 allow for 'rand' at prompt #2 to randomly select pics in file    "
    Print #FileFile, "new feature 25 august 2002 display elapsed time and character count on display"
    Print #FileFile, "new feature 26 august 2002 CC search uses 'S' search then regets with a 'P' previous sequence"
    Print #FileFile, "new feature 31 August 2002 if search data numeric don't do uppercase if alpha switch to uppercase for search"
    Print #FileFile, "new feature 05 October 2002 PAUSE / PA option allow for screen saver type display of text, stopping for TT on each screen"
    Print #FileFile, "new feature 12 October 2002 DEBPHO to debug the photo display on some computers"
    Print #FileFile, "new feature 30 November 2002 DISC now shows os version ie Windows 98 etc"
    Print #FileFile, "option change 05 December 2002 P1 & P2 just change image size P does single photo display"
    Print #FileFile, "new feature 09 December 2002 Cmd(49) RANDOM / NORANDOM store from run to run"
    Print #FileFile, "new feature 24 December 2002 Cmd(50) showfiles / noshowfiles in merged text files show file name where match found"
    Print #FileFile, "option change if program name matches Cmd(45) then change drive # to that of app.path"
    Print #FileFile, "Allow close of window during screen saver to unload form"
    Print #FileFile, "new feature 01 February 2003 allow for *.avi video play cmd(51) cmd(52)"
    Print #FileFile, "new feature 28 February 2003 enable skip if same file listed twice for play or display cmd(54)"
    Print #FileFile, "new feature 20 March 2003 context clear cmd(55) allow for centering of text on screen"
    Print #FileFile, "New Feature 22 March 2003 PHOTO_DETAIL text info along with pics see Cmd(56)"
    Print #FileFile, "New Feature 26 March 2003 version 1.03 & 1.02b switchable time display"
    Print #FileFile, "Hint: In random display mode, to view last 3 pictures add 6 blank lines to file"
    Print #FileFile, "New Feature 11 May 2003 ver=1.05 display *.mpg video files similar to *.avi"
    Print #FileFile, "New Feature 24 June 2003 ver=1.07 play mp3 audio files similar to above"
    Print #FileFile, "New Feature 13 July 2003 ver=1.07a ELAP for elapsed time display after mp3 and mpg files"
    Print #FileFile, "New Feature 19 July 2003 ver=1.07T THUMbnail at prompt #2 to just show or play short clips / previews /samples using tt time"
    Print #FileFile, "      Version 2.0 and greater "
    Print #FileFile, "          START==10000 at prompt #2 to test various start points in a video or song"
    Print #FileFile, "          START==20000 on the search detail line, to change individual file start point"
    Print #FileFile, "          use START==????? and WAIT=??? to show / play mini-clips from anywhere in a file"
    Print #FileFile, "          ***start==1000 for those video files that just won't play ie.(wavie lines only)"
    Print #FileFile, "26 July 2003 default context display color to value in cmd(29) alternate color"
    Print #FileFile, "          minor bug fix related to wait time changing"
    Print #FileFile, "New Feature 03 August 2003 ver=2.2 allow for FIT== And REG== to determine regular or fit to screen similar to WAIT="
    Print #FileFile, "New Feature 18 August 2003 THUMB== on detail line will force a short play based on delay seconds"
    Print #FileFile, "                           DOSHOW to over-ride any noshow settings (on detail line)"
    Print #FileFile, "Fix 09 September 2003 to allow for service pack 1 systems"
    Print #FileFile, "New Feature 24 September 2003 freeze=5 on a mpg video, after a wait=10 will hold the frame for display for 5 seconds"
    Print #FileFile, "New Feature 29 September 2003 cmd(59) element used to adjust delay time (between systems)"
    Print #FileFile, "New Feature 30 October 2003 cmd(61) element SETSPEED and speed=500 at prompt #2"
    Print #FileFile, "New Feature 16 November 2003 over-ride above speed on a line basis ie speed=250"
    Print #FileFile, "Fix 08 January 2004 replace SET SPEED with PAUSE and RESUME sequences"
    Print #FileFile, "New Feature 16 August 2004 WMV video format now plays"
    Print #FileFile, "Fix 16 August 2004 pause between jpg files fixed"
    Print #FileFile, "Fix 27 August 2004 at interrupt alow for complete play by entering aa"
    Print #FileFile, "New Feature 04 September allow for multiple inputs on prompt #2 seperated by / sep"
    Print #FileFile, "New Feature 12 September 2004 allow for background job backgrd==c:\search\backgrd1\backgrd1.exe "
    Print #FileFile, "New Feature cmd(68) EOF_STOP to allow for background job to stop at eof use with (WW) prompt #2 option"
    Print #FileFile, "New Feature Cmd(69) HIT_STOP to allow for background job to stop at first match/hit use with (P) prompt #2 option"
    Print #FileFile, "New Feature Cmd(70) FOREGROUND causes a background job to be given focus over main job"
    Print #FileFile, "New Feature command==tt8/rand/randa/thumb/ww found on match line will change command settings"
    Print #FileFile, "Fix 25 October 2004 Small video now play"
    Print #FileFile, "27 October 2004 Batch file input option see cmd(71)"
    Print #FileFile, "    differences between to files limit of 10000 lines for now"
    Print #FileFile, "    differences between to files limit of 20000 lines for now 18-Mar-2007"
    Print #FileFile, "Minor change force directory date out on gf option 03 November 2004"
    Print #FileFile, "08 November 2004 create a results.txt file for video pics music see Cmd(72)"
    Print #FileFile, "18 November 2004 minor New Feature DETAIL, NODETAIL changes cmd(56)"
    Print #FileFile, "19 November 2004 New feature allows for CONTROL==controlx.txt or controly.txt to switch control files"
    Print #FileFile, "     ***vip*** change the control files on the fly (very usefull)"
    Print #FileFile, "     less use for the command== option now"    '19 November 2004
                    '20 November 2004 allow for a cr at photo continue prompt to restart screen saver if auto_exe set
    Print #FileFile, "22 November 2004 New feature on CONTROL== if cmd(73) has FILESWITCH then change the input file to element cmd(46)"
    Print #FileFile, "        in the new control file. "        '22 November 2004
    Print #FileFile, "23 November 2004 New Feature on OPTIONS==OPT=controla.txtOPT=controlb.txtOPT=controlc.txt select 1 2 or 3 ..10 for navigation of presentations"
    Print #FileFile, "28 November 2004 New Feature add GIF files to list of pictures displayed"
    Print #FileFile, "06 December 2004 New Feature enter DX at search prompt for yesterday date but not 01"
    Print #FileFile, "17 December 2004 New Feature enter a V or H at the search prompt will paste in the clipboard just like Ctrl V"
    Print #FileFile, "19 December 2004 New Feature allow for inset keypad at prompt #1 and prompt #3 ie J=4 K=5 L=6 etc"
    Print #FileFile, "24 December 2004 New Feature allow for pause after each print line. Rather than all text displayed at once see Cmd(76)"
    Print #FileFile, "06 January 2005 Minor New Feature, if pause do not display prompt #4 at end of screen add 1 line to max_cnt cmd(1)"
    Print #FileFile, "07 January 2005 New Feature allow for merge of text data files other than .txt ie .htm .for .frm .cls etc..."
    Print #FileFile, "10 January 2005 Minor fix re change 06 January 2005 the F option was not continuing now it does"
    Print #FileFile, "20 August 2005 allow for the last few seconds as a thumbnail using start==299000 and thumb"
    Print #FileFile, "29 November 2006 instead of pictures allow for URL address to be copied and pasted    "
    Print #FileFile, "29 January 2007 anything after http: if no / follows keep all except the http:    "
    Print #FileFile, "21 February 2007 fix the Start== not working properly - just initialize a couple values?"
    Print #FileFile, "11 March 2007 final fix for the start== bug"
    Print #FileFile, "28 March 2007 remove the favorites from directory it causes me problems"
    Print #FileFile, "05 February 2008 fix timer. to speed up scrolling text display    "
    Print #FileFile, "22 November 2010 allow for a negative wait=-3 to play all but the final 3 seconds of a video 22Nov2010  "
    Print #FileFile, "20 November 2011 the BEGIN and AGAIN features were added to video playback in november"
    Print #FileFile, "31 December 2011 for G to be entered on interrupt and have the clip play again just like above"
    Print #FileFile, "06 February 2012 FF for a catalog with fast forward capabilities "
    Print #FileFile, "09 June 2016 minor change for Windows 10 operating system    "
    Print #FileFile, "17 December 2017 allow for video playback to move ahead in small or large increments budge_amt    "
    Print #FileFile, "12 March 2018 initial budge playback feature added more work needed    "
    Print #FileFile, "Version = " + vvvesion
    Print #FileFile, "contact " + program_info + " stonedan@telusplanet.net"
    Close #FileFile
    DoEvents
'check the latest features above
        Return
        
line_30920:         '18 august 2002
    zzz_cnt = 0
'        xtemp = InputBox(" testing doug noonerror " + CStr(zzz_cnt), "testing Prompt   ", , xx1 - offset1, yy1 - offset2)
    On Error GoTo error_line_30920
    OutFile = FreeFile
    Open TheFile For Input As #OutFile
next_line_30920:
    Line Input #OutFile, aaa
    zzz_cnt = zzz_cnt + 1
    If zzz_cnt >= cnt Then GoTo exit_line_30920
    GoTo next_line_30920
error_line_30920:
    Resume resume_line_30920
resume_line_30920:
    On Error GoTo Errors_31000
exit_line_30920:
    Return          '18 august 2002
    
Errors_31000:
    'the following line for debugging only ***vip***
'january 09 2001 below   27 june 2003 testing the stuff on tran's computer protection probably
'    Print " Save-yore error"; Err.Number; " "; save_line; " "; Err.Description; " "; LastFile
'        tt1 = InputBox("error=" + CStr(Err.Number) + " " + save_line, , , 4400, 4500) 'TESTING 09 February 2004
'16Apr2014  trap the end of file on the reverse file read here
'        tt1 = InputBox("Error trap info error=" + CStr(Err.Number) + " " + save_line, , , 4400, 4500) '21Mar2016
    If save_line = "70" Then
        Resume line_74
    End If          '16Apr2014
    If debug_photo Then
        tt1 = InputBox("Error trap routine info error=" + " " + Err.Description + CStr(Err.Number) + " " + save_line, , , 4400, 4500)  '16Apr2014
    End If          '16Apr2014
'    tt1 = InputBox("Error trap =" + UCase(ttt) + CStr(Err.Number) + " " + save_line, , , 4400, 4500) '09Aug2016
'    If Err.Number = 277 And copy_photo = "YES" Then
'        photo_dir = ztemp
'        Resume line_2153 'test only
'    End If      '09Aug2016
    If save_line = "17059b" Then
        frmproj2.Caption = "clipboard loaded "      '09Jun2016 form description
        Resume Photo_continue_prompt
    End If          '09Jun2016  testing
    If save_line = "29310" Then
        Cls
        Print "bad data record???"
        Print Left(aaa, 100) '17dec2017
        Resume line_29399
    End If          '17Dec2017 was getting an error on weird line length call
    If save_line = "29400" Then
        Resume line_29450
    End If          '17Dec2017 was getting an error on weird line length call
        
    
    If save_line = "30711b" Then
        Resume line_30730
    End If                  '05 March 2010
    
    If save_line = "testing" Then
        Resume End_32000
    End If                  'december 15 2000
    If save_line = "16100-8d" And Err.Number = 20 Then
        frmproj2.Caption = " skiping folder " + cript1(JJ) '01 march 2010
        Resume line_30711
    End If                  '03 March 2010
    If save_line = "17200" Then
        Resume line_17220
    End If                  'february 24 2002
    If save_line = "30610" Then
        Resume line_30640
    End If                  'january 09 2001
    If save_line = "50" Then
        Resume What_50
    End If                  'december 6 2000
    If Err.Number = 62 And save_line = "28010" Then
        Close FileFile
        DoEvents
'        Kill "C:\control.txt"
        Kill "control.txt"      'december 3 2000
        FileFile = FreeFile
        II = DoEvents
'        Open "C:\control.txt" For Output Access Write As #FileFile
        Open "control.txt" For Output Access Write As #FileFile 'december 3 2000
        For f = 1 To 100
            If Len(Cmd(f)) < 1 Then
                Cmd(f) = Space(50)
            End If
            Print #FileFile, Cmd(f)
        Next f
        Close FileFile
        II = DoEvents
'        tt1 = InputBox("extending control.txt", , , 4400, 4500)  'TESTING ONLY
        Resume line_20
     
    End If
'===============================
  
    If Err.Number = 62 And save_line = "1000" And _
        end_cnt < 10 Then
        end_cnt = end_cnt + 1
        Resume input_1000
    End If      'allow for more than 1 eof marker in file
    If Err.Number = 62 And save_line = "29010" Then
        Resume line_29090
    End If
    If save_line = "29092" Then
        Resume line_29093
    End If
    If Err.Number = 53 And save_line = "29005" Then
        Resume line_29008
    End If

'"Y" logic below read till end of file then back off a few lines
'    then simulate a "C" and "A" for all search.
    If Err.Number = 62 And save_line = "65" Then
            Close #OutFile
            DoEvents
            OutFile = FreeFile
            Open TheFile For Input As #OutFile
            DoEvents
'            tt1 = InputBox("testing only " + TheFile, , , 4400, 4500) 'TESTING ONLY

        For bbb = 1 To zzz_cnt - MAX_CNT + 2    '2 TO SEE THE LAST 2 STATS LINES
            Line Input #OutFile, aaa
        Next bbb
        Previous_line = aaa
        Line Input #OutFile, aaa
        zzz_cnt = bbb
        search_prompt = ""  'this keeps the next prompt#3 from being "D"
        prompt2 = "C"
        SAVE_ttt = "C"
        inin = "A"
        SSS1 = "A"
        ttt = "C"
        printed_cnt = 0         'this fixes the 2 execution of this function
                                'ie the lines print from the top not midway
        Resume input_1000a
    End If      'december 09 2001
    
'05 october 2002    If Err.Number = 62 And save_line = "1000" And sscreen_saver = "Y" Then
'20 November 2004 maybe the save_line thing is a problem when I return to save_line = 1000a etc
            frmproj2.Caption = " eofsw test=" + CStr(Err.Number) + "=" + save_line + " " + Err.Description + " " + Pict_file '26 November 2004
'14Apr2012    If Err.Number = 62 And save_line = "1000" And (sscreen_saver = "Y" Or text_pause) Then
    If Err.Number = 62 And save_line = "1000" And (sscreen_saver = "Y" Or text_pause) Then
'14Apr2012 used to test for error that was fixed  below      tt1 = InputBox("at problem area   " + CStr(zzz_cnt) + " " + CStr(rand_no), , , 4400, 4500) 'TESTING ONLY
        If rand = -1 Then
           Randomize
        frmproj2.Caption = "29Oct2012aaa testing doug randomizer rnd" + CStr(rand_no) + " " + CStr(zzz_cnt)
          
          rand_cnt = zzz_cnt - 3
          zzz_cnt = 0
           rand_no = Int(rand_cnt * Rnd + 1)
           rand_str = "random#" + CStr(rand_no) '18Mar2012
            Resume Do_Search_110 '14Apr2012
        End If                  '14Apr2012
'=====================================================================================
    If Left(Cmd(74), 8) = "EOFCMD==" Then
        control_file = Trim(Right(Cmd(74), Len(Cmd(74)) - 8)) '26 November 2004
        aaa = "CONTROL==" + Right(Cmd(74), Len(Cmd(74)) - 8)
'            tempdata = Cmd(73)                      '22 November 2004
        temptemp = Trim(Cmd(48))        '26 November 2004
        tempdata = "CONTROL"            '26 November 2004
        
        eofsw = "YES"                   '26 November 2004
        GoSub Control_28000     'change the control file from in a *.txt file
        Resume eof_entry                '26 November 2004
'        GoTo eof_entry                '27 November 2004
    End If                              '26 November 2004
    
   If Left(UCase(Cmd(68)), 8) = "EOF_STOP" Then GoTo End_32000 '12 September 2004
'====================================================================================
        
        Close #OutFile
'    frmproj2.Caption = program_info + " (cls #17)" '18Dec2013
        Cls                     '20 August 2003
'        Print "at end of file prompt"  '04 april 2002
'        Print "at end of file prompt"  '10 January 2010 skip this for special program derulaswamp
       If Left(UCase(Cmd(80)), 13) = "PROMPTDETAILS" Then Print "at end of file prompt " + sscreen_saver '

'        Print "at end of file prompt"  '04 april 2002
    '16 April 2004 make it so that after the first complete pass that the logic
    '              will switch so that the data is randomly accessed. This is only
    '              for auto-run custom cd's
    
'20 November 2004 found this may be a problem ***vip*** todo check this once and a while.
'        If save_line <> "1000" Then
        If save_line = "1000111" Then
'===================================
        rand = -1
        Randomize
        frmproj2.Caption = "29Oct2012yyy testing doug randomizer rnd" + CStr(rand_no) + " " + CStr(zzz_cnt)
        
        rand_cnt = zzz_cnt - 3
        zzz_cnt = 0
        rand_no = Int(rand_cnt * Rnd + 1)
        rand_str = "random#" + CStr(rand_no) '18Mar2012
        Cmd(49) = "RANDOM"
        
        rand1 = -1
        Cmd(65) = "RANDBEG"

        thumb_nail = "YES"
        Cmd(67) = "THUMB"
 
        play_speed = 250
        save_play_speed = 250
        Cmd(61) = "SETSPEED250"
 '===================================
        
        End If              '16 April 2004 deactivate this code if not a auto-run cd
        
        Beep
        play_speed = 1000      '13 May 2004 the first one after end of file was playing slomo this fixed
        hold_sec = Val(Cmd(27))    '29 april 2002
'        Cmd(27) = Format(delay_sec, "###0.0000") '01 may 2002
'25 March 2004        delay_sec = 1   '04 april 2002
'25 March 2004        GoSub line_30000    '04 april 2002
        new_delay_sec = 1   '25 March 2004
'        new_delay_sec = 10   '20 November 2004
'         frmproj2.Caption = " end of file dougdoug " + interrupt_prompt2 + "*" + SSS1 + "*" + SSS2 '20 November 2004 test
        If interrupt_prompt2 = "WW" Then
            sscreen_saver_ww = "YES" '20 November 2004 test
            sscreen_saver = "Y"     '20 November 2004 test
            prompt2 = "WW"      'this one seems to keep it going where before it stopped. now it gets wrong data
            prompt2 = "SS"      '20 November 2004
            tt1 = "WW"
   
