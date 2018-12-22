VERSION 5.00
Begin VB.Form Form1
  Caption = "Registration Demo"
  BackColor = &HFFFFFF&
  ScaleMode = 3
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 1 'Fixed Single
  Icon = "Form1.frx":0000
  LinkTopic = "Form1"
  MaxButton = 0   'False
  MinButton = 0   'False
  ClientLeft = -15
  ClientTop = 375
  ClientWidth = 9270
  ClientHeight = 4320
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame Frame3
    Caption = "3. Masukkan No. Lisensi untuk mendapatkan Full Version"
    BackColor = &HFFFFFF&
    Left = 240
    Top = 2760
    Width = 8775
    Height = 1335
    TabIndex = 6
    BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8,25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.CommandButton cmdRegister
      Caption = "Register"
      Left = 3600
      Top = 720
      Width = 1095
      Height = 285
      TabIndex = 9
    End
    Begin VB.TextBox txtGetRegKey
      Left = 1560
      Top = 360
      Width = 1935
      Height = 285
      Text = "Input No Lisensi disini"
      TabIndex = 8
    End
    Begin VB.TextBox txtValidationStatus
      Left = 1560
      Top = 720
      Width = 1935
      Height = 285
      Enabled = 0   'False
      Text = "ValidationStatus"
      TabIndex = 7
    End
    Begin VB.Label Label5
      Caption = "No. Lisensi"
      BackColor = &HFFFFFF&
      Left = 120
      Top = 360
      Width = 1335
      Height = 255
      TabIndex = 15
      Alignment = 1 'Right Justify
    End
    Begin VB.Label Label3
      Caption = "Registration Status"
      BackColor = &HFFFFFF&
      Left = 120
      Top = 750
      Width = 1335
      Height = 255
      TabIndex = 11
      Alignment = 1 'Right Justify
    End
  End
  Begin VB.Frame Frame2
    Caption = "2. Masukkan ID kedalam kotak, setelah itu sms ke 085731230992 untuk mendapatkan No. lisensi"
    BackColor = &HFFFFFF&
    Left = 240
    Top = 1320
    Width = 8775
    Height = 1215
    TabIndex = 2
    BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8,25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.CommandButton cmdCopyRegKey
      Caption = "Copy"
      Left = 3600
      Top = 720
      Width = 1095
      Height = 285
      Visible = 0   'False
      TabIndex = 5
    End
    Begin VB.TextBox txtGetId
      Left = 1560
      Top = 360
      Width = 1935
      Height = 285
      Visible = 0   'False
      Text = "Paste ID komputer disini"
      TabIndex = 4
    End
    Begin VB.TextBox txtRegKey
      Left = 1560
      Top = 720
      Width = 1935
      Height = 285
      Visible = 0   'False
      Text = "RegKey"
      TabIndex = 3
    End
    Begin VB.Label Label4
      Caption = "ID Komputer"
      BackColor = &HFFFFFF&
      Left = 240
      Top = 360
      Width = 1215
      Height = 255
      TabIndex = 14
      Alignment = 1 'Right Justify
    End
    Begin VB.Label Label2
      Caption = "Registration Key"
      BackColor = &HFFFFFF&
      Left = 240
      Top = 750
      Width = 1215
      Height = 255
      Visible = 0   'False
      TabIndex = 10
      Alignment = 1 'Right Justify
    End
  End
  Begin VB.Frame Frame1
    Caption = "1. ID Komputer Khusus untuk  lisensi program"
    BackColor = &HFFFFFF&
    Left = 240
    Top = 240
    Width = 8775
    Height = 855
    TabIndex = 0
    BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8,25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.TextBox txtId
      Left = 1560
      Top = 360
      Width = 1935
      Height = 285
      Text = "MachineId"
      TabIndex = 12
    End
    Begin VB.CommandButton cmdCopyId
      Caption = "Copy"
      Left = 3600
      Top = 360
      Width = 1095
      Height = 285
      TabIndex = 1
    End
    Begin VB.Label Label1
      Caption = "ID Komputer"
      BackColor = &HFFFFFF&
      Left = 480
      Top = 390
      Width = 975
      Height = 255
      TabIndex = 13
      Alignment = 1 'Right Justify
    End
  End
End

Attribute VB_Name = "Form1"


Private Sub txtGetId_Change() '661460
  loc_00661460: push ebp
  loc_00661461: mov ebp, esp
  loc_00661463: sub esp, 0000000Ch
  loc_00661466: push 00405696h ; __vbaExceptHandler
  loc_0066146B: mov eax, fs:[00000000h]
  loc_00661471: push eax
  loc_00661472: mov fs:[00000000h], esp
  loc_00661479: sub esp, 000000DCh
  loc_0066147F: push ebx
  loc_00661480: push esi
  loc_00661481: push edi
  loc_00661482: mov var_C, esp
  loc_00661485: mov var_8, 004054E8h
  loc_0066148C: mov esi, Me
  loc_0066148F: mov eax, esi
  loc_00661491: and eax, 00000001h
  loc_00661494: mov var_4, eax
  loc_00661497: and esi, FFFFFFFEh
  loc_0066149A: push esi
  loc_0066149B: mov Me, esi
  loc_0066149E: mov ecx, [esi]
  loc_006614A0: call [ecx+00000004h]
  loc_006614A3: mov ebx, [00401130h] ; __vbaAryConstruct2
  loc_006614A9: push 00000002h
  loc_006614AB: lea edx, var_34
  loc_006614AE: xor edi, edi
  loc_006614B0: push 00438BD4h
  loc_006614B5: push edx
  loc_006614B6: mov var_1C, edi
  loc_006614B9: mov var_58, edi
  loc_006614BC: mov var_5C, edi
  loc_006614BF: mov var_60, edi
  loc_006614C2: mov var_64, edi
  loc_006614C5: mov var_68, edi
  loc_006614C8: mov var_6C, edi
  loc_006614CB: mov var_7C, edi
  loc_006614CE: mov var_8C, edi
  loc_006614D4: mov var_9C, edi
  loc_006614DA: mov var_B0, edi
  loc_006614E0: mov var_B4, edi
  loc_006614E6: call ebx
  loc_006614E8: push 00000008h
  loc_006614EA: lea eax, var_50
  loc_006614ED: push 0043FA14h
  loc_006614F2: push eax
  loc_006614F3: call ebx
  loc_006614F5: mov ecx, [esi]
  loc_006614F7: push esi
  loc_006614F8: call [ecx+00000730h]
  loc_006614FE: mov edx, [esi]
  loc_00661500: push esi
  loc_00661501: call [edx+0000031Ch]
  loc_00661507: push eax
  loc_00661508: lea eax, var_68
  loc_0066150B: push eax
  loc_0066150C: call [004010B8h] ; __vbaObjSet
  loc_00661512: mov ebx, eax
  loc_00661514: lea edx, var_58
  loc_00661517: push edx
  loc_00661518: push ebx
  loc_00661519: mov ecx, [ebx]
  loc_0066151B: call [ecx+000000A0h]
  loc_00661521: cmp eax, edi
  loc_00661523: fnclex
  loc_00661525: jge 00661539h
  loc_00661527: push 000000A0h
  loc_0066152C: push 0042DFCCh
  loc_00661531: push ebx
  loc_00661532: push eax
  loc_00661533: call [00401080h] ; __vbaHresultCheckObj
  loc_00661539: mov eax, var_58
  loc_0066153C: push eax
  loc_0066153D: push 0042DFECh
  loc_00661542: call [0040112Ch] ; __vbaStrCmp
  loc_00661548: mov ebx, eax
  loc_0066154A: lea ecx, var_58
  loc_0066154D: neg ebx
  loc_0066154F: sbb ebx, ebx
  loc_00661551: inc ebx
  loc_00661552: neg ebx
  loc_00661554: call [0040129Ch] ; __vbaFreeStr
  loc_0066155A: lea ecx, var_68
  loc_0066155D: call [004012A0h] ; __vbaFreeObj
  loc_00661563: cmp bx, di
  loc_00661566: jnz 00661AD0h
  loc_0066156C: lea ecx, var_7C
  loc_0066156F: mov var_74, 80020004h
  loc_00661576: push ecx
  loc_00661577: mov var_7C, 0000000Ah
  loc_0066157E: call [004010A8h] ; rtcRandomize
  loc_00661584: lea ecx, var_7C
  loc_00661587: call [00401020h] ; __vbaFreeVar
  loc_0066158D: mov ebx, 00000001h
  loc_00661592: mov eax, 00000002h
  loc_00661597: cmp bx, ax
  loc_0066159A: jg 006616B1h
  loc_006615A0: lea edx, var_7C
  loc_006615A3: mov var_74, 80020004h
  loc_006615AA: push edx
  loc_006615AB: mov var_7C, 0000000Ah
  loc_006615B2: call [0040109Ch] ; rtcRandomNext
  loc_006615B8: fstp real4 ptr var_B0
  loc_006615BE: movsx edi, bx
  loc_006615C1: cmp edi, 00000003h
  loc_006615C4: jb 006615CCh
  loc_006615C6: call [00401120h] ; __vbaGenerateBoundsError
  loc_006615CC: fld real4 ptr var_B0
  loc_006615D2: fmul st0, real4 ptr [004054C4h]
  loc_006615D8: fadd st0, real4 ptr [004054C0h]
  loc_006615DE: fnstsw ax
  loc_006615E0: test al, 0Dh
  loc_006615E2: jnz 00661B71h
  loc_006615E8: call [00401258h] ; __vbaR8IntI2
  loc_006615EE: mov ecx, var_28
  loc_006615F1: mov [ecx+edi*2], ax
  loc_006615F5: lea ecx, var_7C
  loc_006615F8: call [00401020h] ; __vbaFreeVar
  loc_006615FE: cmp edi, 00000003h
  loc_00661601: mov var_74, 0000000Bh
  loc_00661608: mov var_7C, 00000002h
  loc_0066160F: jb 00661617h
  loc_00661611: call [00401120h] ; __vbaGenerateBoundsError
  loc_00661617: mov ecx, var_28
  loc_0066161A: lea edx, [esi+0000003Ch]
  loc_0066161D: mov var_94, edx
  loc_00661623: mov var_9C, 00004008h
  loc_0066162D: movsx edx, [ecx+edi*2]
  loc_00661631: lea eax, var_7C
  loc_00661634: lea ecx, var_8C
  loc_0066163A: push eax
  loc_0066163B: lea eax, var_9C
  loc_00661641: push edx
  loc_00661642: push eax
  loc_00661643: push ecx
  loc_00661644: call [00401100h] ; rtcMidCharVar
  loc_0066164A: cmp edi, 00000003h
  loc_0066164D: jb 00661655h
  loc_0066164F: call [00401120h] ; __vbaGenerateBoundsError
  loc_00661655: lea edx, var_8C
  loc_0066165B: push edx
  loc_0066165C: call [00401034h] ; __vbaStrVarMove
  loc_00661662: mov edx, eax
  loc_00661664: lea ecx, var_58
  loc_00661667: call [0040126Ch] ; __vbaStrMove
  loc_0066166D: mov edx, eax
  loc_0066166F: mov eax, var_44
  loc_00661672: lea ecx, [eax+edi*4]
  loc_00661675: call [004011FCh] ; __vbaStrCopy
  loc_0066167B: lea ecx, var_58
  loc_0066167E: call [0040129Ch] ; __vbaFreeStr
  loc_00661684: lea ecx, var_8C
  loc_0066168A: lea edx, var_7C
  loc_0066168D: push ecx
  loc_0066168E: push edx
  loc_0066168F: push 00000002h
  loc_00661691: call [0040103Ch] ; __vbaFreeVarList
  loc_00661697: mov eax, 00000001h
  loc_0066169C: add esp, 0000000Ch
  loc_0066169F: add ax, bx
  loc_006616A2: jo 00661B76h
  loc_006616A8: mov ebx, eax
  loc_006616AA: xor edi, edi
  loc_006616AC: jmp 00661592h
  loc_006616B1: mov eax, var_44
  loc_006616B4: mov ecx, [eax+00000004h]
  loc_006616B7: push ecx
  loc_006616B8: call [004012A4h] ; rtcR8ValFromBstr
  loc_006616BE: mov edx, var_44
  loc_006616C1: fstp real8 ptr var_BC
  loc_006616C7: mov eax, [edx+00000008h]
  loc_006616CA: push eax
  loc_006616CB: call [004012A4h] ; rtcR8ValFromBstr
  loc_006616D1: mov ecx, [esi]
  loc_006616D3: push esi
  loc_006616D4: fstp real8 ptr var_C4
  loc_006616DA: call [ecx+0000031Ch]
  loc_006616E0: lea edx, var_68
  loc_006616E3: push eax
  loc_006616E4: push edx
  loc_006616E5: call [004010B8h] ; __vbaObjSet
  loc_006616EB: mov ebx, eax
  loc_006616ED: lea ecx, var_58
  loc_006616F0: push ecx
  loc_006616F1: push ebx
  loc_006616F2: mov eax, [ebx]
  loc_006616F4: call [eax+000000A0h]
  loc_006616FA: cmp eax, edi
  loc_006616FC: fnclex
  loc_006616FE: jge 00661712h
  loc_00661700: push 000000A0h
  loc_00661705: push 0042DFCCh
  loc_0066170A: push ebx
  loc_0066170B: push eax
  loc_0066170C: call [00401080h] ; __vbaHresultCheckObj
  loc_00661712: mov edx, var_58
  loc_00661715: push edx
  loc_00661716: call [004012A4h] ; rtcR8ValFromBstr
  loc_0066171C: fsub st0, real8 ptr var_BC
  loc_00661722: lea ecx, var_8C
  loc_00661728: mov var_7C, 00000005h
  loc_0066172F: fadd st0, real8 ptr var_C4
  loc_00661735: fstp real8 ptr var_74
  loc_00661738: fnstsw ax
  loc_0066173A: test al, 0Dh
  loc_0066173C: jnz 00661B71h
  loc_00661742: lea eax, var_7C
  loc_00661745: push eax
  loc_00661746: push ecx
  loc_00661747: call [00401240h] ; rtcVarStrFromVar
  loc_0066174D: lea edx, var_8C
  loc_00661753: push edx
  loc_00661754: call [00401034h] ; __vbaStrVarMove
  loc_0066175A: mov ebx, [0040126Ch] ; __vbaStrMove
  loc_00661760: mov edx, eax
  loc_00661762: lea ecx, var_1C
  loc_00661765: call ebx
  loc_00661767: lea ecx, var_58
  loc_0066176A: call [0040129Ch] ; __vbaFreeStr
  loc_00661770: lea ecx, var_68
  loc_00661773: call [004012A0h] ; __vbaFreeObj
  loc_00661779: lea eax, var_8C
  loc_0066177F: lea ecx, var_7C
  loc_00661782: push eax
  loc_00661783: push ecx
  loc_00661784: push 00000002h
  loc_00661786: call [0040103Ch] ; __vbaFreeVarList
  loc_0066178C: mov edx, var_28
  loc_0066178F: mov edi, [0040100Ch] ; __vbaStrI2
  loc_00661795: add esp, 0000000Ch
  loc_00661798: mov ax, [edx+00000002h]
  loc_0066179C: push eax
  loc_0066179D: call edi
  loc_0066179F: mov edx, eax
  loc_006617A1: lea ecx, var_58
  loc_006617A4: call ebx
  loc_006617A6: mov ecx, var_28
  loc_006617A9: push eax
  loc_006617AA: mov dx, [ecx+00000004h]
  loc_006617AE: push edx
  loc_006617AF: call edi
  loc_006617B1: mov edx, eax
  loc_006617B3: lea ecx, var_5C
  loc_006617B6: call ebx
  loc_006617B8: mov edi, [00401070h] ; __vbaStrCat
  loc_006617BE: push eax
  loc_006617BF: call edi
  loc_006617C1: mov edx, eax
  loc_006617C3: lea ecx, var_60
  loc_006617C6: call ebx
  loc_006617C8: push eax
  loc_006617C9: mov eax, var_1C
  loc_006617CC: push 00000000h
  loc_006617CE: push FFFFFFFFh
  loc_006617D0: push 00000001h
  loc_006617D2: push 0042DFECh
  loc_006617D7: push 0043FDECh
  loc_006617DC: push eax
  loc_006617DD: call [00401194h] ; rtcReplace
  loc_006617E3: mov edx, eax
  loc_006617E5: lea ecx, var_64
  loc_006617E8: call ebx
  loc_006617EA: push eax
  loc_006617EB: call edi
  loc_006617ED: mov edx, eax
  loc_006617EF: lea ecx, var_1C
  loc_006617F2: call ebx
  loc_006617F4: lea ecx, var_64
  loc_006617F7: lea edx, var_60
  loc_006617FA: push ecx
  loc_006617FB: lea eax, var_5C
  loc_006617FE: push edx
  loc_006617FF: lea ecx, var_58
  loc_00661802: push eax
  loc_00661803: push ecx
  loc_00661804: push 00000004h
  loc_00661806: call [00401204h] ; __vbaFreeStrList
  loc_0066180C: mov edx, [esi]
  loc_0066180E: add esp, 00000014h
  loc_00661811: push esi
  loc_00661812: call [edx+00000320h]
  loc_00661818: push eax
  loc_00661819: lea eax, var_68
  loc_0066181C: push eax
  loc_0066181D: call [004010B8h] ; __vbaObjSet
  loc_00661823: mov edx, var_1C
  loc_00661826: mov edi, eax
  loc_00661828: push edx
  loc_00661829: push edi
  loc_0066182A: mov ecx, [edi]
  loc_0066182C: call [ecx+000000A4h]
  loc_00661832: test eax, eax
  loc_00661834: fnclex
  loc_00661836: jge 0066184Ah
  loc_00661838: push 000000A4h
  loc_0066183D: push 0042DFCCh
  loc_00661842: push edi
  loc_00661843: push eax
  loc_00661844: call [00401080h] ; __vbaHresultCheckObj
  loc_0066184A: lea ecx, var_68
  loc_0066184D: call [004012A0h] ; __vbaFreeObj
  loc_00661853: mov eax, [esi]
  loc_00661855: push esi
  loc_00661856: call [eax+00000320h]
  loc_0066185C: mov ebx, [004010B8h] ; __vbaObjSet
  loc_00661862: lea ecx, var_68
  loc_00661865: push eax
  loc_00661866: push ecx
  loc_00661867: call ebx
  loc_00661869: mov edi, eax
  loc_0066186B: lea eax, var_58
  loc_0066186E: push eax
  loc_0066186F: push edi
  loc_00661870: mov edx, [edi]
  loc_00661872: call [edx+000000A0h]
  loc_00661878: test eax, eax
  loc_0066187A: fnclex
  loc_0066187C: jge 00661890h
  loc_0066187E: push 000000A0h
  loc_00661883: push 0042DFCCh
  loc_00661888: push edi
  loc_00661889: push eax
  loc_0066188A: call [00401080h] ; __vbaHresultCheckObj
  loc_00661890: mov ecx, var_58
  loc_00661893: push ecx
  loc_00661894: call [00401030h] ; __vbaLenBstr
  loc_0066189A: mov ecx, eax
  loc_0066189C: call [00401138h] ; __vbaI2I4
  loc_006618A2: lea ecx, var_58
  loc_006618A5: mov var_18, eax
  loc_006618A8: call [0040129Ch] ; __vbaFreeStr
  loc_006618AE: lea ecx, var_68
  loc_006618B1: call [004012A0h] ; __vbaFreeObj
  loc_006618B7: mov ecx, var_18
  loc_006618BA: mov eax, 00000001h
  loc_006618BF: cmp cx, ax
  loc_006618C2: jl 00661A90h
  loc_006618C8: movsx eax, cx
  loc_006618CB: cmp eax, 00000005h
  loc_006618CE: jz 006618DEh
  loc_006618D0: cmp eax, 0000000Ah
  loc_006618D3: jz 006618DEh
  loc_006618D5: cmp eax, 0000000Fh
  loc_006618D8: jnz 00661A7Bh
  loc_006618DE: mov edx, [esi]
  loc_006618E0: push esi
  loc_006618E1: call [edx+00000320h]
  loc_006618E7: push eax
  loc_006618E8: lea eax, var_68
  loc_006618EB: push eax
  loc_006618EC: call ebx
  loc_006618EE: mov edi, eax
  loc_006618F0: movsx eax, var_18
  loc_006618F4: mov ecx, [edi]
  loc_006618F6: push eax
  loc_006618F7: push edi
  loc_006618F8: call [ecx+00000114h]
  loc_006618FE: test eax, eax
  loc_00661900: fnclex
  loc_00661902: jge 00661916h
  loc_00661904: push 00000114h
  loc_00661909: push 0042DFCCh
  loc_0066190E: push edi
  loc_0066190F: push eax
  loc_00661910: call [00401080h] ; __vbaHresultCheckObj
  loc_00661916: lea ecx, var_68
  loc_00661919: call [004012A0h] ; __vbaFreeObj
  loc_0066191F: mov edx, [esi]
  loc_00661921: push esi
  loc_00661922: call [edx+00000320h]
  loc_00661928: push eax
  loc_00661929: lea eax, var_68
  loc_0066192C: push eax
  loc_0066192D: call ebx
  loc_0066192F: mov edi, eax
  loc_00661931: push 00000000h
  loc_00661933: push edi
  loc_00661934: mov ecx, [edi]
  loc_00661936: call [ecx+0000011Ch]
  loc_0066193C: test eax, eax
  loc_0066193E: fnclex
  loc_00661940: jge 00661954h
  loc_00661942: push 0000011Ch
  loc_00661947: push 0042DFCCh
  loc_0066194C: push edi
  loc_0066194D: push eax
  loc_0066194E: call [00401080h] ; __vbaHresultCheckObj
  loc_00661954: lea ecx, var_68
  loc_00661957: call [004012A0h] ; __vbaFreeObj
  loc_0066195D: mov edx, [esi]
  loc_0066195F: push esi
  loc_00661960: call [edx+00000320h]
  loc_00661966: push eax
  loc_00661967: lea eax, var_68
  loc_0066196A: push eax
  loc_0066196B: call ebx
  loc_0066196D: mov edi, eax
  loc_0066196F: lea edx, var_58
  loc_00661972: push edx
  loc_00661973: push edi
  loc_00661974: mov ecx, [edi]
  loc_00661976: call [ecx+000000A0h]
  loc_0066197C: test eax, eax
  loc_0066197E: fnclex
  loc_00661980: jge 00661994h
  loc_00661982: push 000000A0h
  loc_00661987: push 0042DFCCh
  loc_0066198C: push edi
  loc_0066198D: push eax
  loc_0066198E: call [00401080h] ; __vbaHresultCheckObj
  loc_00661994: mov eax, var_58
  loc_00661997: push eax
  loc_00661998: call [00401030h] ; __vbaLenBstr
  loc_0066199E: movsx edi, var_18
  loc_006619A2: xor ecx, ecx
  loc_006619A4: cmp edi, eax
  loc_006619A6: setl cl
  loc_006619A9: neg ecx
  loc_006619AB: mov di, cx
  loc_006619AE: lea ecx, var_58
  loc_006619B1: call [0040129Ch] ; __vbaFreeStr
  loc_006619B7: lea ecx, var_68
  loc_006619BA: call [004012A0h] ; __vbaFreeObj
  loc_006619C0: test di, di
  loc_006619C3: jz 00661A7Bh
  loc_006619C9: mov edx, [esi]
  loc_006619CB: push esi
  loc_006619CC: call [edx+00000320h]
  loc_006619D2: push eax
  loc_006619D3: lea eax, var_6C
  loc_006619D6: push eax
  loc_006619D7: call ebx
  loc_006619D9: mov ecx, [esi]
  loc_006619DB: push esi
  loc_006619DC: mov ebx, eax
  loc_006619DE: call [ecx+00000320h]
  loc_006619E4: lea edx, var_68
  loc_006619E7: push eax
  loc_006619E8: push edx
  loc_006619E9: call [004010B8h] ; __vbaObjSet
  loc_006619EF: mov edi, eax
  loc_006619F1: lea ecx, var_58
  loc_006619F4: push ecx
  loc_006619F5: push edi
  loc_006619F6: mov eax, [edi]
  loc_006619F8: call [eax+00000120h]
  loc_006619FE: test eax, eax
  loc_00661A00: fnclex
  loc_00661A02: jge 00661A16h
  loc_00661A04: push 00000120h
  loc_00661A09: push 0042DFCCh
  loc_00661A0E: push edi
  loc_00661A0F: push eax
  loc_00661A10: call [00401080h] ; __vbaHresultCheckObj
  loc_00661A16: mov edx, var_58
  loc_00661A19: mov edi, [ebx]
  loc_00661A1B: push edx
  loc_00661A1C: push 0042F350h ; "-"
  loc_00661A21: call [00401070h] ; __vbaStrCat
  loc_00661A27: mov edx, eax
  loc_00661A29: lea ecx, var_5C
  loc_00661A2C: call [0040126Ch] ; __vbaStrMove
  loc_00661A32: push eax
  loc_00661A33: push ebx
  loc_00661A34: call [edi+00000124h]
  loc_00661A3A: test eax, eax
  loc_00661A3C: fnclex
  loc_00661A3E: jge 00661A52h
  loc_00661A40: push 00000124h
  loc_00661A45: push 0042DFCCh
  loc_00661A4A: push ebx
  loc_00661A4B: push eax
  loc_00661A4C: call [00401080h] ; __vbaHresultCheckObj
  loc_00661A52: lea eax, var_5C
  loc_00661A55: lea ecx, var_58
  loc_00661A58: push eax
  loc_00661A59: push ecx
  loc_00661A5A: push 00000002h
  loc_00661A5C: call [00401204h] ; __vbaFreeStrList
  loc_00661A62: lea edx, var_6C
  loc_00661A65: lea eax, var_68
  loc_00661A68: push edx
  loc_00661A69: push eax
  loc_00661A6A: push 00000002h
  loc_00661A6C: call [0040104Ch] ; __vbaFreeObjList
  loc_00661A72: mov ebx, [004010B8h] ; __vbaObjSet
  loc_00661A78: add esp, 00000018h
  loc_00661A7B: or eax, FFFFFFFFh
  loc_00661A7E: add ax, var_18
  loc_00661A82: jo 00661B76h
  loc_00661A88: mov var_18, eax
  loc_00661A8B: jmp 006618B7h
  loc_00661A90: mov ecx, [esi]
  loc_00661A92: push esi
  loc_00661A93: call [ecx+00000318h]
  loc_00661A99: lea edx, var_68
  loc_00661A9C: push eax
  loc_00661A9D: push edx
  loc_00661A9E: call ebx
  loc_00661AA0: mov esi, eax
  loc_00661AA2: push FFFFFFFFh
  loc_00661AA4: push esi
  loc_00661AA5: mov eax, [esi]
  loc_00661AA7: call [eax+0000008Ch]
  loc_00661AAD: test eax, eax
  loc_00661AAF: fnclex
  loc_00661AB1: jge 00661AC5h
  loc_00661AB3: push 0000008Ch
  loc_00661AB8: push 0043FC70h
  loc_00661ABD: push esi
  loc_00661ABE: push eax
  loc_00661ABF: call [00401080h] ; __vbaHresultCheckObj
  loc_00661AC5: lea ecx, var_68
  loc_00661AC8: call [004012A0h] ; __vbaFreeObj
  loc_00661ACE: xor edi, edi
  loc_00661AD0: mov var_4, edi
  loc_00661AD3: fwait
  loc_00661AD4: push 00661B52h
  loc_00661AD9: jmp 00661B1Ah
  loc_00661ADB: lea ecx, var_64
  loc_00661ADE: lea edx, var_60
  loc_00661AE1: push ecx
  loc_00661AE2: lea eax, var_5C
  loc_00661AE5: push edx
  loc_00661AE6: lea ecx, var_58
  loc_00661AE9: push eax
  loc_00661AEA: push ecx
  loc_00661AEB: push 00000004h
  loc_00661AED: call [00401204h] ; __vbaFreeStrList
  loc_00661AF3: lea edx, var_6C
  loc_00661AF6: lea eax, var_68
  loc_00661AF9: push edx
  loc_00661AFA: push eax
  loc_00661AFB: push 00000002h
  loc_00661AFD: call [0040104Ch] ; __vbaFreeObjList
  loc_00661B03: lea ecx, var_8C
  loc_00661B09: lea edx, var_7C
  loc_00661B0C: push ecx
  loc_00661B0D: push edx
  loc_00661B0E: push 00000002h
  loc_00661B10: call [0040103Ch] ; __vbaFreeVarList
  loc_00661B16: add esp, 0000002Ch
  loc_00661B19: ret
  loc_00661B1A: lea ecx, var_1C
  loc_00661B1D: call [0040129Ch] ; __vbaFreeStr
  loc_00661B23: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00661B29: lea ecx, var_B0
  loc_00661B2F: lea eax, var_34
  loc_00661B32: push ecx
  loc_00661B33: push 00000000h
  loc_00661B35: mov var_B0, eax
  loc_00661B3B: call __vbaAryDestruct
  loc_00661B3D: lea eax, var_B4
  loc_00661B43: lea edx, var_50
  loc_00661B46: push eax
  loc_00661B47: push 00000000h
  loc_00661B49: mov var_B4, edx
  loc_00661B4F: call __vbaAryDestruct
  loc_00661B51: ret
  loc_00661B52: mov eax, Me
  loc_00661B55: push eax
  loc_00661B56: mov ecx, [eax]
  loc_00661B58: call [ecx+00000008h]
  loc_00661B5B: mov eax, var_4
  loc_00661B5E: mov ecx, var_14
  loc_00661B61: pop edi
  loc_00661B62: pop esi
  loc_00661B63: mov fs:[00000000h], ecx
  loc_00661B6A: pop ebx
  loc_00661B6B: mov esp, ebp
  loc_00661B6D: pop ebp
  loc_00661B6E: retn 0004h
End Sub

Private Sub txtGetId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '661B80
  loc_00661B80: push ebp
  loc_00661B81: mov ebp, esp
  loc_00661B83: sub esp, 0000000Ch
  loc_00661B86: push 00405696h ; __vbaExceptHandler
  loc_00661B8B: mov eax, fs:[00000000h]
  loc_00661B91: push eax
  loc_00661B92: mov fs:[00000000h], esp
  loc_00661B99: sub esp, 00000024h
  loc_00661B9C: push ebx
  loc_00661B9D: push esi
  loc_00661B9E: push edi
  loc_00661B9F: mov var_C, esp
  loc_00661BA2: mov var_8, 004054F8h
  loc_00661BA9: mov esi, Me
  loc_00661BAC: mov eax, esi
  loc_00661BAE: and eax, 00000001h
  loc_00661BB1: mov var_4, eax
  loc_00661BB4: and esi, FFFFFFFEh
  loc_00661BB7: push esi
  loc_00661BB8: mov Me, esi
  loc_00661BBB: mov ecx, [esi]
  loc_00661BBD: call [ecx+00000004h]
  loc_00661BC0: mov edx, [esi]
  loc_00661BC2: xor eax, eax
  loc_00661BC4: push esi
  loc_00661BC5: mov var_18, eax
  loc_00661BC8: mov var_1C, eax
  loc_00661BCB: mov var_20, eax
  loc_00661BCE: call [edx+0000031Ch]
  loc_00661BD4: mov ebx, [004010B8h] ; __vbaObjSet
  loc_00661BDA: push eax
  loc_00661BDB: lea eax, var_1C
  loc_00661BDE: push eax
  loc_00661BDF: call ebx
  loc_00661BE1: mov edi, eax
  loc_00661BE3: push 00000000h
  loc_00661BE5: push edi
  loc_00661BE6: mov ecx, [edi]
  loc_00661BE8: call [ecx+00000114h]
  loc_00661BEE: test eax, eax
  loc_00661BF0: fnclex
  loc_00661BF2: jge 00661C06h
  loc_00661BF4: push 00000114h
  loc_00661BF9: push 0042DFCCh
  loc_00661BFE: push edi
  loc_00661BFF: push eax
  loc_00661C00: call [00401080h] ; __vbaHresultCheckObj
  loc_00661C06: lea ecx, var_1C
  loc_00661C09: call [004012A0h] ; __vbaFreeObj
  loc_00661C0F: mov edx, [esi]
  loc_00661C11: push esi
  loc_00661C12: call [edx+0000031Ch]
  loc_00661C18: push eax
  loc_00661C19: lea eax, var_20
  loc_00661C1C: push eax
  loc_00661C1D: call ebx
  loc_00661C1F: mov ecx, [esi]
  loc_00661C21: push esi
  loc_00661C22: mov edi, eax
  loc_00661C24: call [ecx+0000031Ch]
  loc_00661C2A: lea edx, var_1C
  loc_00661C2D: push eax
  loc_00661C2E: push edx
  loc_00661C2F: call ebx
  loc_00661C31: mov esi, eax
  loc_00661C33: lea ecx, var_18
  loc_00661C36: push ecx
  loc_00661C37: push esi
  loc_00661C38: mov eax, [esi]
  loc_00661C3A: call [eax+000000A0h]
  loc_00661C40: test eax, eax
  loc_00661C42: fnclex
  loc_00661C44: jge 00661C58h
  loc_00661C46: push 000000A0h
  loc_00661C4B: push 0042DFCCh
  loc_00661C50: push esi
  loc_00661C51: push eax
  loc_00661C52: call [00401080h] ; __vbaHresultCheckObj
  loc_00661C58: mov edx, var_18
  loc_00661C5B: mov esi, [edi]
  loc_00661C5D: push edx
  loc_00661C5E: call [00401030h] ; __vbaLenBstr
  loc_00661C64: push eax
  loc_00661C65: push edi
  loc_00661C66: call [esi+0000011Ch]
  loc_00661C6C: test eax, eax
  loc_00661C6E: fnclex
  loc_00661C70: jge 00661C84h
  loc_00661C72: push 0000011Ch
  loc_00661C77: push 0042DFCCh
  loc_00661C7C: push edi
  loc_00661C7D: push eax
  loc_00661C7E: call [00401080h] ; __vbaHresultCheckObj
  loc_00661C84: lea ecx, var_18
  loc_00661C87: call [0040129Ch] ; __vbaFreeStr
  loc_00661C8D: lea eax, var_20
  loc_00661C90: lea ecx, var_1C
  loc_00661C93: push eax
  loc_00661C94: push ecx
  loc_00661C95: push 00000002h
  loc_00661C97: call [0040104Ch] ; __vbaFreeObjList
  loc_00661C9D: add esp, 0000000Ch
  loc_00661CA0: mov var_4, 00000000h
  loc_00661CA7: push 00661CCCh
  loc_00661CAC: jmp 00661CCBh
  loc_00661CAE: lea ecx, var_18
  loc_00661CB1: call [0040129Ch] ; __vbaFreeStr
  loc_00661CB7: lea edx, var_20
  loc_00661CBA: lea eax, var_1C
  loc_00661CBD: push edx
  loc_00661CBE: push eax
  loc_00661CBF: push 00000002h
  loc_00661CC1: call [0040104Ch] ; __vbaFreeObjList
  loc_00661CC7: add esp, 0000000Ch
  loc_00661CCA: ret
  loc_00661CCB: ret
  loc_00661CCC: mov eax, Me
  loc_00661CCF: push eax
  loc_00661CD0: mov ecx, [eax]
  loc_00661CD2: call [ecx+00000008h]
  loc_00661CD5: mov eax, var_4
  loc_00661CD8: mov ecx, var_14
  loc_00661CDB: pop edi
  loc_00661CDC: pop esi
  loc_00661CDD: mov fs:[00000000h], ecx
  loc_00661CE4: pop ebx
  loc_00661CE5: mov esp, ebp
  loc_00661CE7: pop ebp
  loc_00661CE8: retn 0014h
End Sub

Private Sub txtGetRegKey_Change() '661CF0
  loc_00661CF0: push ebp
  loc_00661CF1: mov ebp, esp
  loc_00661CF3: sub esp, 0000000Ch
  loc_00661CF6: push 00405696h ; __vbaExceptHandler
  loc_00661CFB: mov eax, fs:[00000000h]
  loc_00661D01: push eax
  loc_00661D02: mov fs:[00000000h], esp
  loc_00661D09: sub esp, 00000014h
  loc_00661D0C: push ebx
  loc_00661D0D: push esi
  loc_00661D0E: push edi
  loc_00661D0F: mov var_C, esp
  loc_00661D12: mov var_8, 00405508h
  loc_00661D19: mov esi, Me
  loc_00661D1C: mov eax, esi
  loc_00661D1E: and eax, 00000001h
  loc_00661D21: mov var_4, eax
  loc_00661D24: and esi, FFFFFFFEh
  loc_00661D27: push esi
  loc_00661D28: mov Me, esi
  loc_00661D2B: mov ecx, [esi]
  loc_00661D2D: call [ecx+00000004h]
  loc_00661D30: mov edx, [esi]
  loc_00661D32: xor edi, edi
  loc_00661D34: push esi
  loc_00661D35: mov var_18, edi
  loc_00661D38: call [edx+00000300h]
  loc_00661D3E: push eax
  loc_00661D3F: lea eax, var_18
  loc_00661D42: push eax
  loc_00661D43: call [004010B8h] ; __vbaObjSet
  loc_00661D49: mov esi, eax
  loc_00661D4B: push FFFFFFFFh
  loc_00661D4D: push esi
  loc_00661D4E: mov ecx, [esi]
  loc_00661D50: call [ecx+0000008Ch]
  loc_00661D56: cmp eax, edi
  loc_00661D58: fnclex
  loc_00661D5A: jge 00661D6Eh
  loc_00661D5C: push 0000008Ch
  loc_00661D61: push 0043FC70h
  loc_00661D66: push esi
  loc_00661D67: push eax
  loc_00661D68: call [00401080h] ; __vbaHresultCheckObj
  loc_00661D6E: lea ecx, var_18
  loc_00661D71: call [004012A0h] ; __vbaFreeObj
  loc_00661D77: mov var_4, edi
  loc_00661D7A: push 00661D8Ch
  loc_00661D7F: jmp 00661D8Bh
  loc_00661D81: lea ecx, var_18
  loc_00661D84: call [004012A0h] ; __vbaFreeObj
  loc_00661D8A: ret
  loc_00661D8B: ret
  loc_00661D8C: mov eax, Me
  loc_00661D8F: push eax
  loc_00661D90: mov edx, [eax]
  loc_00661D92: call [edx+00000008h]
  loc_00661D95: mov eax, var_4
  loc_00661D98: mov ecx, var_14
  loc_00661D9B: pop edi
  loc_00661D9C: pop esi
  loc_00661D9D: mov fs:[00000000h], ecx
  loc_00661DA4: pop ebx
  loc_00661DA5: mov esp, ebp
  loc_00661DA7: pop ebp
  loc_00661DA8: retn 0004h
End Sub

Private Sub txtGetRegKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '661DB0
  loc_00661DB0: push ebp
  loc_00661DB1: mov ebp, esp
  loc_00661DB3: sub esp, 0000000Ch
  loc_00661DB6: push 00405696h ; __vbaExceptHandler
  loc_00661DBB: mov eax, fs:[00000000h]
  loc_00661DC1: push eax
  loc_00661DC2: mov fs:[00000000h], esp
  loc_00661DC9: sub esp, 00000024h
  loc_00661DCC: push ebx
  loc_00661DCD: push esi
  loc_00661DCE: push edi
  loc_00661DCF: mov var_C, esp
  loc_00661DD2: mov var_8, 00405518h
  loc_00661DD9: mov esi, Me
  loc_00661DDC: mov eax, esi
  loc_00661DDE: and eax, 00000001h
  loc_00661DE1: mov var_4, eax
  loc_00661DE4: and esi, FFFFFFFEh
  loc_00661DE7: push esi
  loc_00661DE8: mov Me, esi
  loc_00661DEB: mov ecx, [esi]
  loc_00661DED: call [ecx+00000004h]
  loc_00661DF0: mov edx, [esi]
  loc_00661DF2: xor eax, eax
  loc_00661DF4: push esi
  loc_00661DF5: mov var_18, eax
  loc_00661DF8: mov var_1C, eax
  loc_00661DFB: mov var_20, eax
  loc_00661DFE: call [edx+00000304h]
  loc_00661E04: mov ebx, [004010B8h] ; __vbaObjSet
  loc_00661E0A: push eax
  loc_00661E0B: lea eax, var_1C
  loc_00661E0E: push eax
  loc_00661E0F: call ebx
  loc_00661E11: mov edi, eax
  loc_00661E13: push 00000000h
  loc_00661E15: push edi
  loc_00661E16: mov ecx, [edi]
  loc_00661E18: call [ecx+00000114h]
  loc_00661E1E: test eax, eax
  loc_00661E20: fnclex
  loc_00661E22: jge 00661E36h
  loc_00661E24: push 00000114h
  loc_00661E29: push 0042DFCCh
  loc_00661E2E: push edi
  loc_00661E2F: push eax
  loc_00661E30: call [00401080h] ; __vbaHresultCheckObj
  loc_00661E36: lea ecx, var_1C
  loc_00661E39: call [004012A0h] ; __vbaFreeObj
  loc_00661E3F: mov edx, [esi]
  loc_00661E41: push esi
  loc_00661E42: call [edx+00000304h]
  loc_00661E48: push eax
  loc_00661E49: lea eax, var_20
  loc_00661E4C: push eax
  loc_00661E4D: call ebx
  loc_00661E4F: mov ecx, [esi]
  loc_00661E51: push esi
  loc_00661E52: mov edi, eax
  loc_00661E54: call [ecx+00000304h]
  loc_00661E5A: lea edx, var_1C
  loc_00661E5D: push eax
  loc_00661E5E: push edx
  loc_00661E5F: call ebx
  loc_00661E61: mov esi, eax
  loc_00661E63: lea ecx, var_18
  loc_00661E66: push ecx
  loc_00661E67: push esi
  loc_00661E68: mov eax, [esi]
  loc_00661E6A: call [eax+000000A0h]
  loc_00661E70: test eax, eax
  loc_00661E72: fnclex
  loc_00661E74: jge 00661E88h
  loc_00661E76: push 000000A0h
  loc_00661E7B: push 0042DFCCh
  loc_00661E80: push esi
  loc_00661E81: push eax
  loc_00661E82: call [00401080h] ; __vbaHresultCheckObj
  loc_00661E88: mov edx, var_18
  loc_00661E8B: mov esi, [edi]
  loc_00661E8D: push edx
  loc_00661E8E: call [00401030h] ; __vbaLenBstr
  loc_00661E94: push eax
  loc_00661E95: push edi
  loc_00661E96: call [esi+0000011Ch]
  loc_00661E9C: test eax, eax
  loc_00661E9E: fnclex
  loc_00661EA0: jge 00661EB4h
  loc_00661EA2: push 0000011Ch
  loc_00661EA7: push 0042DFCCh
  loc_00661EAC: push edi
  loc_00661EAD: push eax
  loc_00661EAE: call [00401080h] ; __vbaHresultCheckObj
  loc_00661EB4: lea ecx, var_18
  loc_00661EB7: call [0040129Ch] ; __vbaFreeStr
  loc_00661EBD: lea eax, var_20
  loc_00661EC0: lea ecx, var_1C
  loc_00661EC3: push eax
  loc_00661EC4: push ecx
  loc_00661EC5: push 00000002h
  loc_00661EC7: call [0040104Ch] ; __vbaFreeObjList
  loc_00661ECD: add esp, 0000000Ch
  loc_00661ED0: mov var_4, 00000000h
  loc_00661ED7: push 00661EFCh
  loc_00661EDC: jmp 00661EFBh
  loc_00661EDE: lea ecx, var_18
  loc_00661EE1: call [0040129Ch] ; __vbaFreeStr
  loc_00661EE7: lea edx, var_20
  loc_00661EEA: lea eax, var_1C
  loc_00661EED: push edx
  loc_00661EEE: push eax
  loc_00661EEF: push 00000002h
  loc_00661EF1: call [0040104Ch] ; __vbaFreeObjList
  loc_00661EF7: add esp, 0000000Ch
  loc_00661EFA: ret
  loc_00661EFB: ret
  loc_00661EFC: mov eax, Me
  loc_00661EFF: push eax
  loc_00661F00: mov ecx, [eax]
  loc_00661F02: call [ecx+00000008h]
  loc_00661F05: mov eax, var_4
  loc_00661F08: mov ecx, var_14
  loc_00661F0B: pop edi
  loc_00661F0C: pop esi
  loc_00661F0D: mov fs:[00000000h], ecx
  loc_00661F14: pop ebx
  loc_00661F15: mov esp, ebp
  loc_00661F17: pop ebp
  loc_00661F18: retn 0014h
End Sub

Private Sub cmdCopyRegKey_Click() '65F090
  loc_0065F090: push ebp
  loc_0065F091: mov ebp, esp
  loc_0065F093: sub esp, 0000000Ch
  loc_0065F096: push 00405696h ; __vbaExceptHandler
  loc_0065F09B: mov eax, fs:[00000000h]
  loc_0065F0A1: push eax
  loc_0065F0A2: mov fs:[00000000h], esp
  loc_0065F0A9: sub esp, 0000003Ch
  loc_0065F0AC: push ebx
  loc_0065F0AD: push esi
  loc_0065F0AE: push edi
  loc_0065F0AF: mov var_C, esp
  loc_0065F0B2: mov var_8, 00405490h
  loc_0065F0B9: mov edi, Me
  loc_0065F0BC: mov eax, edi
  loc_0065F0BE: and eax, 00000001h
  loc_0065F0C1: mov var_4, eax
  loc_0065F0C4: and edi, FFFFFFFEh
  loc_0065F0C7: push edi
  loc_0065F0C8: mov Me, edi
  loc_0065F0CB: mov ecx, [edi]
  loc_0065F0CD: call [ecx+00000004h]
  loc_0065F0D0: mov eax, [0066A318h]
  loc_0065F0D5: xor ebx, ebx
  loc_0065F0D7: cmp eax, ebx
  loc_0065F0D9: mov var_18, ebx
  loc_0065F0DC: mov var_1C, ebx
  loc_0065F0DF: mov var_20, ebx
  loc_0065F0E2: jnz 0065F0F4h
  loc_0065F0E4: push 0066A318h
  loc_0065F0E9: push 0042E390h
  loc_0065F0EE: call [004011E8h] ; __vbaNew2
  loc_0065F0F4: mov esi, [0066A318h]
  loc_0065F0FA: lea eax, var_20
  loc_0065F0FD: push eax
  loc_0065F0FE: push esi
  loc_0065F0FF: mov edx, [esi]
  loc_0065F101: call [edx+0000001Ch]
  loc_0065F104: cmp eax, ebx
  loc_0065F106: fnclex
  loc_0065F108: jge 0065F119h
  loc_0065F10A: push 0000001Ch
  loc_0065F10C: push 0042E380h
  loc_0065F111: push esi
  loc_0065F112: push eax
  loc_0065F113: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F119: mov ecx, [edi]
  loc_0065F11B: mov ebx, var_20
  loc_0065F11E: push edi
  loc_0065F11F: call [ecx+00000320h]
  loc_0065F125: lea edx, var_1C
  loc_0065F128: push eax
  loc_0065F129: push edx
  loc_0065F12A: call [004010B8h] ; __vbaObjSet
  loc_0065F130: mov esi, eax
  loc_0065F132: lea ecx, var_18
  loc_0065F135: push ecx
  loc_0065F136: push esi
  loc_0065F137: mov eax, [esi]
  loc_0065F139: call [eax+000000A0h]
  loc_0065F13F: test eax, eax
  loc_0065F141: fnclex
  loc_0065F143: jge 0065F157h
  loc_0065F145: push 000000A0h
  loc_0065F14A: push 0042DFCCh
  loc_0065F14F: push esi
  loc_0065F150: push eax
  loc_0065F151: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F157: sub esp, 00000010h
  loc_0065F15A: mov eax, 0000000Ah
  loc_0065F15F: mov ecx, esp
  loc_0065F161: mov edx, [ebx]
  loc_0065F163: mov [ecx], eax
  loc_0065F165: mov eax, var_2C
  loc_0065F168: mov [ecx+00000004h], eax
  loc_0065F16B: mov eax, 80020004h
  loc_0065F170: mov [ecx+00000008h], eax
  loc_0065F173: mov eax, var_24
  loc_0065F176: mov [ecx+0000000Ch], eax
  loc_0065F179: mov ecx, var_18
  loc_0065F17C: push ecx
  loc_0065F17D: push ebx
  loc_0065F17E: call [edx+00000060h]
  loc_0065F181: test eax, eax
  loc_0065F183: fnclex
  loc_0065F185: jge 0065F196h
  loc_0065F187: push 00000060h
  loc_0065F189: push 0043FA48h
  loc_0065F18E: push ebx
  loc_0065F18F: push eax
  loc_0065F190: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F196: lea ecx, var_18
  loc_0065F199: call [0040129Ch] ; __vbaFreeStr
  loc_0065F19F: lea edx, var_20
  loc_0065F1A2: lea eax, var_1C
  loc_0065F1A5: push edx
  loc_0065F1A6: push eax
  loc_0065F1A7: push 00000002h
  loc_0065F1A9: call [0040104Ch] ; __vbaFreeObjList
  loc_0065F1AF: mov ecx, [edi]
  loc_0065F1B1: add esp, 0000000Ch
  loc_0065F1B4: push edi
  loc_0065F1B5: call [ecx+00000304h]
  loc_0065F1BB: lea edx, var_1C
  loc_0065F1BE: push eax
  loc_0065F1BF: push edx
  loc_0065F1C0: call [004010B8h] ; __vbaObjSet
  loc_0065F1C6: mov esi, eax
  loc_0065F1C8: push FFFFFFFFh
  loc_0065F1CA: push esi
  loc_0065F1CB: mov eax, [esi]
  loc_0065F1CD: call [eax+00000094h]
  loc_0065F1D3: test eax, eax
  loc_0065F1D5: fnclex
  loc_0065F1D7: jge 0065F1EBh
  loc_0065F1D9: push 00000094h
  loc_0065F1DE: push 0042DFCCh
  loc_0065F1E3: push esi
  loc_0065F1E4: push eax
  loc_0065F1E5: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F1EB: lea ecx, var_1C
  loc_0065F1EE: call [004012A0h] ; __vbaFreeObj
  loc_0065F1F4: mov var_4, 00000000h
  loc_0065F1FB: push 0065F220h
  loc_0065F200: jmp 0065F21Fh
  loc_0065F202: lea ecx, var_18
  loc_0065F205: call [0040129Ch] ; __vbaFreeStr
  loc_0065F20B: lea ecx, var_20
  loc_0065F20E: lea edx, var_1C
  loc_0065F211: push ecx
  loc_0065F212: push edx
  loc_0065F213: push 00000002h
  loc_0065F215: call [0040104Ch] ; __vbaFreeObjList
  loc_0065F21B: add esp, 0000000Ch
  loc_0065F21E: ret
  loc_0065F21F: ret
  loc_0065F220: mov eax, Me
  loc_0065F223: push eax
  loc_0065F224: mov ecx, [eax]
  loc_0065F226: call [ecx+00000008h]
  loc_0065F229: mov eax, var_4
  loc_0065F22C: mov ecx, var_14
  loc_0065F22F: pop edi
  loc_0065F230: pop esi
  loc_0065F231: mov fs:[00000000h], ecx
  loc_0065F238: pop ebx
  loc_0065F239: mov esp, ebp
  loc_0065F23B: pop ebp
  loc_0065F23C: retn 0004h
End Sub

Private Sub Form_Load() '65F950
  loc_0065F950: push ebp
  loc_0065F951: mov ebp, esp
  loc_0065F953: sub esp, 0000000Ch
  loc_0065F956: push 00405696h ; __vbaExceptHandler
  loc_0065F95B: mov eax, fs:[00000000h]
  loc_0065F961: push eax
  loc_0065F962: mov fs:[00000000h], esp
  loc_0065F969: sub esp, 000001B8h
  loc_0065F96F: push ebx
  loc_0065F970: push esi
  loc_0065F971: push edi
  loc_0065F972: mov var_C, esp
  loc_0065F975: mov var_8, 004054B0h
  loc_0065F97C: mov esi, Me
  loc_0065F97F: mov eax, esi
  loc_0065F981: and eax, 00000001h
  loc_0065F984: mov var_4, eax
  loc_0065F987: and esi, FFFFFFFEh
  loc_0065F98A: push esi
  loc_0065F98B: mov Me, esi
  loc_0065F98E: mov ecx, [esi]
  loc_0065F990: call [ecx+00000004h]
  loc_0065F993: push 00000008h
  loc_0065F995: lea edx, var_50
  loc_0065F998: xor edi, edi
  loc_0065F99A: push 0043FD44h
  loc_0065F99F: push edx
  loc_0065F9A0: mov var_24, edi
  loc_0065F9A3: mov var_34, edi
  loc_0065F9A6: mov var_38, edi
  loc_0065F9A9: mov var_58, edi
  loc_0065F9AC: mov var_60, edi
  loc_0065F9AF: mov var_64, edi
  loc_0065F9B2: mov var_68, edi
  loc_0065F9B5: mov var_6C, edi
  loc_0065F9B8: mov var_70, edi
  loc_0065F9BB: mov var_74, edi
  loc_0065F9BE: mov var_84, edi
  loc_0065F9C4: mov var_94, edi
  loc_0065F9CA: mov var_A4, edi
  loc_0065F9D0: mov var_B4, edi
  loc_0065F9D6: mov var_C4, edi
  loc_0065F9DC: mov var_D4, edi
  loc_0065F9E2: mov var_E4, edi
  loc_0065F9E8: mov var_F4, edi
  loc_0065F9EE: mov var_104, edi
  loc_0065F9F4: mov var_118, edi
  loc_0065F9FA: mov var_144, edi
  loc_0065FA00: mov var_154, edi
  loc_0065FA06: mov var_164, edi
  loc_0065FA0C: mov var_174, edi
  loc_0065FA12: mov var_184, edi
  loc_0065FA18: mov var_194, edi
  loc_0065FA1E: mov var_1A4, edi
  loc_0065FA24: mov var_1B4, edi
  loc_0065FA2A: call [00401130h] ; __vbaAryConstruct2
  loc_0065FA30: mov eax, [esi]
  loc_0065FA32: push esi
  loc_0065FA33: call [eax+00000700h]
  loc_0065FA39: cmp eax, edi
  loc_0065FA3B: jge 0065FA4Fh
  loc_0065FA3D: push 00000700h
  loc_0065FA42: push 00434D5Ch
  loc_0065FA47: push esi
  loc_0065FA48: push eax
  loc_0065FA49: call [00401080h] ; __vbaHresultCheckObj
  loc_0065FA4F: mov ebx, [004011FCh] ; __vbaStrCopy
  loc_0065FA55: add esi, 0000003Ch
  loc_0065FA58: mov edx, 0043FB08h ; "320420352196268181894210437267790267909582698624445212800344381963372357131848839301613255938768148682177489084666712402256650210971"
  loc_0065FA5D: mov ecx, esi
  loc_0065FA5F: call ebx
  loc_0065FA61: mov edx, [esi]
  loc_0065FA63: mov ecx, var_44
  loc_0065FA66: call ebx
  loc_0065FA68: lea ecx, var_84
  loc_0065FA6E: lea edx, var_D4
  loc_0065FA74: push ecx
  loc_0065FA75: push 00000018h
  loc_0065FA77: lea eax, var_94
  loc_0065FA7D: mov ebx, 00000002h
  loc_0065FA82: push edx
  loc_0065FA83: push eax
  loc_0065FA84: mov var_7C, 00000010h
  loc_0065FA8B: mov var_84, ebx
  loc_0065FA91: mov var_CC, esi
  loc_0065FA97: mov var_D4, 00004008h
  loc_0065FAA1: call [00401100h] ; rtcMidCharVar
  loc_0065FAA7: lea ecx, var_94
  loc_0065FAAD: push ecx
  loc_0065FAAE: call [00401034h] ; __vbaStrVarMove
  loc_0065FAB4: mov edx, eax
  loc_0065FAB6: lea ecx, var_58
  loc_0065FAB9: call [0040126Ch] ; __vbaStrMove
  loc_0065FABF: lea edx, var_94
  loc_0065FAC5: lea eax, var_84
  loc_0065FACB: push edx
  loc_0065FACC: push eax
  loc_0065FACD: push ebx
  loc_0065FACE: call [0040103Ch] ; __vbaFreeVarList
  loc_0065FAD4: mov ecx, var_58
  loc_0065FAD7: add esp, 0000000Ch
  loc_0065FADA: mov esi, 00000001h
  loc_0065FADF: push edi
  loc_0065FAE0: push FFFFFFFFh
  loc_0065FAE2: push esi
  loc_0065FAE3: push 00434A94h
  loc_0065FAE8: push 0042E3A4h
  loc_0065FAED: push ecx
  loc_0065FAEE: call [00401194h] ; rtcReplace
  loc_0065FAF4: mov edx, eax
  loc_0065FAF6: lea ecx, var_58
  loc_0065FAF9: call [0040126Ch] ; __vbaStrMove
  loc_0065FAFF: lea edx, var_D4
  loc_0065FB05: lea eax, var_E4
  loc_0065FB0B: push edx
  loc_0065FB0C: lea ecx, var_F4
  loc_0065FB12: push eax
  loc_0065FB13: lea edx, var_154
  loc_0065FB19: push ecx
  loc_0065FB1A: lea eax, var_144
  loc_0065FB20: push edx
  loc_0065FB21: lea ecx, var_24
  loc_0065FB24: push eax
  loc_0065FB25: push ecx
  loc_0065FB26: mov var_CC, esi
  loc_0065FB2C: mov var_D4, ebx
  loc_0065FB32: mov var_DC, 00000010h
  loc_0065FB3C: mov var_E4, ebx
  loc_0065FB42: mov var_EC, esi
  loc_0065FB48: mov var_F4, ebx
  loc_0065FB4E: call [004010A0h] ; __vbaVarForInit
  loc_0065FB54: mov esi, [0040121Ch] ; __vbaI4Var
  loc_0065FB5A: test eax, eax
  loc_0065FB5C: jz 0065FC7Fh
  loc_0065FB62: lea eax, var_84
  loc_0065FB68: lea ecx, var_24
  loc_0065FB6B: lea edx, var_58
  loc_0065FB6E: push eax
  loc_0065FB6F: push ecx
  loc_0065FB70: mov var_7C, 00000001h
  loc_0065FB77: mov var_84, ebx
  loc_0065FB7D: mov var_CC, edx
  loc_0065FB83: mov var_D4, 00004008h
  loc_0065FB8D: call __vbaI4Var
  loc_0065FB8F: push eax
  loc_0065FB90: lea edx, var_D4
  loc_0065FB96: lea eax, var_94
  loc_0065FB9C: push edx
  loc_0065FB9D: push eax
  loc_0065FB9E: call [00401100h] ; rtcMidCharVar
  loc_0065FBA4: mov ecx, Me
  loc_0065FBA7: lea edx, var_A4
  loc_0065FBAD: push edx
  loc_0065FBAE: mov var_9C, 00000078h
  loc_0065FBB8: lea eax, [ecx+0000003Ch]
  loc_0065FBBB: mov var_A4, ebx
  loc_0065FBC1: mov var_EC, eax
  loc_0065FBC7: lea eax, var_94
  loc_0065FBCD: push eax
  loc_0065FBCE: mov var_F4, 00004008h
  loc_0065FBD8: call __vbaI4Var
  loc_0065FBDA: lea ecx, var_F4
  loc_0065FBE0: push eax
  loc_0065FBE1: lea edx, var_B4
  loc_0065FBE7: push ecx
  loc_0065FBE8: push edx
  loc_0065FBE9: call [00401100h] ; rtcMidCharVar
  loc_0065FBEF: lea eax, var_24
  loc_0065FBF2: push eax
  loc_0065FBF3: call __vbaI4Var
  loc_0065FBF5: cmp eax, 00000011h
  loc_0065FBF8: mov var_124, eax
  loc_0065FBFE: jb 0065FC06h
  loc_0065FC00: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065FC06: lea ecx, var_B4
  loc_0065FC0C: push ecx
  loc_0065FC0D: call [00401034h] ; __vbaStrVarMove
  loc_0065FC13: mov edx, eax
  loc_0065FC15: lea ecx, var_60
  loc_0065FC18: call [0040126Ch] ; __vbaStrMove
  loc_0065FC1E: mov ecx, var_124
  loc_0065FC24: mov edx, eax
  loc_0065FC26: mov eax, var_44
  loc_0065FC29: lea ecx, [eax+ecx*4]
  loc_0065FC2C: call [004011FCh] ; __vbaStrCopy
  loc_0065FC32: lea ecx, var_60
  loc_0065FC35: call [0040129Ch] ; __vbaFreeStr
  loc_0065FC3B: lea edx, var_B4
  loc_0065FC41: lea eax, var_A4
  loc_0065FC47: push edx
  loc_0065FC48: lea ecx, var_94
  loc_0065FC4E: push eax
  loc_0065FC4F: lea edx, var_84
  loc_0065FC55: push ecx
  loc_0065FC56: push edx
  loc_0065FC57: push 00000004h
  loc_0065FC59: call [0040103Ch] ; __vbaFreeVarList
  loc_0065FC5F: add esp, 00000014h
  loc_0065FC62: lea eax, var_154
  loc_0065FC68: lea ecx, var_144
  loc_0065FC6E: lea edx, var_24
  loc_0065FC71: push eax
  loc_0065FC72: push ecx
  loc_0065FC73: push edx
  loc_0065FC74: call [00401290h] ; __vbaVarForNext
  loc_0065FC7A: jmp 0065FB5Ah
  loc_0065FC7F: mov eax, 00000001h
  loc_0065FC84: lea ecx, var_E4
  loc_0065FC8A: mov var_CC, eax
  loc_0065FC90: mov var_EC, eax
  loc_0065FC96: lea eax, var_D4
  loc_0065FC9C: lea edx, var_F4
  loc_0065FCA2: push eax
  loc_0065FCA3: push ecx
  loc_0065FCA4: lea eax, var_174
  loc_0065FCAA: push edx
  loc_0065FCAB: lea ecx, var_164
  loc_0065FCB1: push eax
  loc_0065FCB2: lea edx, var_24
  loc_0065FCB5: push ecx
  loc_0065FCB6: push edx
  loc_0065FCB7: mov var_D4, ebx
  loc_0065FCBD: mov var_DC, 00000078h
  loc_0065FCC7: mov var_E4, ebx
  loc_0065FCCD: mov var_F4, ebx
  loc_0065FCD3: call [004010A0h] ; __vbaVarForInit
  loc_0065FCD9: test eax, eax
  loc_0065FCDB: jz 0065FF21h
  loc_0065FCE1: lea eax, var_D4
  loc_0065FCE7: lea ecx, var_E4
  loc_0065FCED: push eax
  loc_0065FCEE: lea edx, var_F4
  loc_0065FCF4: push ecx
  loc_0065FCF5: lea eax, var_194
  loc_0065FCFB: push edx
  loc_0065FCFC: lea ecx, var_184
  loc_0065FD02: push eax
  loc_0065FD03: lea edx, var_34
  loc_0065FD06: push ecx
  loc_0065FD07: push edx
  loc_0065FD08: mov var_CC, 00000001h
  loc_0065FD12: mov var_D4, ebx
  loc_0065FD18: mov var_DC, 0000000Fh
  loc_0065FD22: mov var_E4, ebx
  loc_0065FD28: mov var_EC, 00000000h
  loc_0065FD32: mov var_F4, ebx
  loc_0065FD38: call [004010A0h] ; __vbaVarForInit
  loc_0065FD3E: test eax, eax
  loc_0065FD40: jz 0065FED3h
  loc_0065FD46: mov eax, 00000001h
  loc_0065FD4B: lea ecx, var_F4
  loc_0065FD51: mov var_AC, eax
  loc_0065FD57: mov var_EC, eax
  loc_0065FD5D: lea eax, var_34
  loc_0065FD60: lea edx, var_A4
  loc_0065FD66: push eax
  loc_0065FD67: push ecx
  loc_0065FD68: push edx
  loc_0065FD69: mov var_B4, ebx
  loc_0065FD6F: mov var_F4, ebx
  loc_0065FD75: call [0040122Ch] ; __vbaVarAdd
  loc_0065FD7B: push eax
  loc_0065FD7C: call __vbaI4Var
  loc_0065FD7E: mov edi, eax
  loc_0065FD80: cmp edi, 00000011h
  loc_0065FD83: jb 0065FD8Bh
  loc_0065FD85: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065FD8B: mov eax, var_44
  loc_0065FD8E: lea edx, var_B4
  loc_0065FD94: push edx
  loc_0065FD95: mov var_104, 00004008h
  loc_0065FD9F: lea ecx, [eax+edi*4]
  loc_0065FDA2: lea eax, var_24
  loc_0065FDA5: push eax
  loc_0065FDA6: mov var_FC, ecx
  loc_0065FDAC: call __vbaI4Var
  loc_0065FDAE: lea ecx, var_104
  loc_0065FDB4: push eax
  loc_0065FDB5: lea edx, var_C4
  loc_0065FDBB: push ecx
  loc_0065FDBC: push edx
  loc_0065FDBD: call [00401100h] ; rtcMidCharVar
  loc_0065FDC3: lea eax, var_C4
  loc_0065FDC9: lea ecx, var_64
  loc_0065FDCC: push eax
  loc_0065FDCD: push ecx
  loc_0065FDCE: call [004011C0h] ; __vbaStrVarVal
  loc_0065FDD4: push eax
  loc_0065FDD5: call [004012A4h] ; rtcR8ValFromBstr
  loc_0065FDDB: fstp real8 ptr var_120
  loc_0065FDE1: lea edx, var_34
  loc_0065FDE4: mov var_7C, 00000001h
  loc_0065FDEB: push edx
  loc_0065FDEC: mov var_84, ebx
  loc_0065FDF2: call __vbaI4Var
  loc_0065FDF4: mov edi, eax
  loc_0065FDF6: cmp edi, 00000011h
  loc_0065FDF9: jb 0065FE01h
  loc_0065FDFB: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065FE01: mov eax, var_44
  loc_0065FE04: lea edx, var_84
  loc_0065FE0A: push edx
  loc_0065FE0B: mov var_D4, 00004008h
  loc_0065FE15: lea ecx, [eax+edi*4]
  loc_0065FE18: lea eax, var_24
  loc_0065FE1B: push eax
  loc_0065FE1C: mov var_CC, ecx
  loc_0065FE22: call __vbaI4Var
  loc_0065FE24: lea ecx, var_D4
  loc_0065FE2A: push eax
  loc_0065FE2B: lea edx, var_94
  loc_0065FE31: push ecx
  loc_0065FE32: push edx
  loc_0065FE33: call [00401100h] ; rtcMidCharVar
  loc_0065FE39: lea eax, var_94
  loc_0065FE3F: lea ecx, var_60
  loc_0065FE42: push eax
  loc_0065FE43: push ecx
  loc_0065FE44: call [004011C0h] ; __vbaStrVarVal
  loc_0065FE4A: push eax
  loc_0065FE4B: call [004012A4h] ; rtcR8ValFromBstr
  loc_0065FE51: fadd st0, real8 ptr var_120
  loc_0065FE57: fnstsw ax
  loc_0065FE59: test al, 0Dh
  loc_0065FE5B: jnz 006603DCh
  loc_0065FE61: call [00401244h] ; __vbaFpI2
  loc_0065FE67: mov edi, eax
  loc_0065FE69: lea edx, var_64
  loc_0065FE6C: lea eax, var_60
  loc_0065FE6F: push edx
  loc_0065FE70: push eax
  loc_0065FE71: push ebx
  loc_0065FE72: call [00401204h] ; __vbaFreeStrList
  loc_0065FE78: lea ecx, var_C4
  loc_0065FE7E: lea edx, var_B4
  loc_0065FE84: push ecx
  loc_0065FE85: lea eax, var_A4
  loc_0065FE8B: push edx
  loc_0065FE8C: lea ecx, var_94
  loc_0065FE92: push eax
  loc_0065FE93: lea edx, var_84
  loc_0065FE99: push ecx
  loc_0065FE9A: push edx
  loc_0065FE9B: push 00000005h
  loc_0065FE9D: call [0040103Ch] ; __vbaFreeVarList
  loc_0065FEA3: add esp, 00000024h
  loc_0065FEA6: cmp di, 0009h
  loc_0065FEAA: jle 0065FEB6h
  loc_0065FEAC: sub di, 000Ah
  loc_0065FEB0: jo 006603E1h
  loc_0065FEB6: lea eax, var_194
  loc_0065FEBC: lea ecx, var_184
  loc_0065FEC2: push eax
  loc_0065FEC3: lea edx, var_34
  loc_0065FEC6: push ecx
  loc_0065FEC7: push edx
  loc_0065FEC8: call [00401290h] ; __vbaVarForNext
  loc_0065FECE: jmp 0065FD3Eh
  loc_0065FED3: mov eax, var_38
  loc_0065FED6: push eax
  loc_0065FED7: push edi
  loc_0065FED8: call [0040100Ch] ; __vbaStrI2
  loc_0065FEDE: mov edx, eax
  loc_0065FEE0: lea ecx, var_60
  loc_0065FEE3: call [0040126Ch] ; __vbaStrMove
  loc_0065FEE9: push eax
  loc_0065FEEA: call [00401070h] ; __vbaStrCat
  loc_0065FEF0: mov edx, eax
  loc_0065FEF2: lea ecx, var_38
  loc_0065FEF5: call [0040126Ch] ; __vbaStrMove
  loc_0065FEFB: lea ecx, var_60
  loc_0065FEFE: call [0040129Ch] ; __vbaFreeStr
  loc_0065FF04: lea ecx, var_174
  loc_0065FF0A: lea edx, var_164
  loc_0065FF10: push ecx
  loc_0065FF11: lea eax, var_24
  loc_0065FF14: push edx
  loc_0065FF15: push eax
  loc_0065FF16: call [00401290h] ; __vbaVarForNext
  loc_0065FF1C: jmp 0065FCD9h
  loc_0065FF21: mov eax, 00000001h
  loc_0065FF26: lea ecx, var_D4
  loc_0065FF2C: mov var_CC, eax
  loc_0065FF32: mov var_EC, eax
  loc_0065FF38: lea edx, var_E4
  loc_0065FF3E: push ecx
  loc_0065FF3F: lea eax, var_F4
  loc_0065FF45: push edx
  loc_0065FF46: lea ecx, var_1B4
  loc_0065FF4C: push eax
  loc_0065FF4D: lea edx, var_1A4
  loc_0065FF53: push ecx
  loc_0065FF54: lea eax, var_24
  loc_0065FF57: push edx
  loc_0065FF58: push eax
  loc_0065FF59: mov var_D4, ebx
  loc_0065FF5F: mov var_DC, 00000010h
  loc_0065FF69: mov var_E4, ebx
  loc_0065FF6F: mov var_F4, ebx
  loc_0065FF75: call [004010A0h] ; __vbaVarForInit
  loc_0065FF7B: test eax, eax
  loc_0065FF7D: jz 0065FFBDh
  loc_0065FF7F: lea ecx, var_24
  loc_0065FF82: push ecx
  loc_0065FF83: call __vbaI4Var
  loc_0065FF85: mov edi, eax
  loc_0065FF87: cmp edi, 00000011h
  loc_0065FF8A: jb 0065FF92h
  loc_0065FF8C: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065FF92: mov eax, var_44
  loc_0065FF95: mov edx, 0042DFECh
  loc_0065FF9A: lea ecx, [eax+edi*4]
  loc_0065FF9D: call [004011FCh] ; __vbaStrCopy
  loc_0065FFA3: lea ecx, var_1B4
  loc_0065FFA9: lea edx, var_1A4
  loc_0065FFAF: push ecx
  loc_0065FFB0: lea eax, var_24
  loc_0065FFB3: push edx
  loc_0065FFB4: push eax
  loc_0065FFB5: call [00401290h] ; __vbaVarForNext
  loc_0065FFBB: jmp 0065FF7Bh
  loc_0065FFBD: mov esi, Me
  loc_0065FFC0: mov ebx, [004011FCh] ; __vbaStrCopy
  loc_0065FFC6: mov edx, 0042DFECh
  loc_0065FFCB: lea edi, [esi+0000003Ch]
  loc_0065FFCE: mov ecx, edi
  loc_0065FFD0: call ebx
  loc_0065FFD2: mov edx, var_38
  loc_0065FFD5: mov ecx, edi
  loc_0065FFD7: call ebx
  loc_0065FFD9: mov eax, [0066A318h]
  loc_0065FFDE: test eax, eax
  loc_0065FFE0: jnz 0065FFF2h
  loc_0065FFE2: push 0066A318h
  loc_0065FFE7: push 0042E390h
  loc_0065FFEC: call [004011E8h] ; __vbaNew2
  loc_0065FFF2: mov edi, [0066A318h]
  loc_0065FFF8: lea edx, var_74
  loc_0065FFFB: push edx
  loc_0065FFFC: push edi
  loc_0065FFFD: mov ecx, [edi]
  loc_0065FFFF: call [ecx+00000014h]
  loc_00660002: test eax, eax
  loc_00660004: fnclex
  loc_00660006: jge 00660017h
  loc_00660008: push 00000014h
  loc_0066000A: push 0042E380h
  loc_0066000F: push edi
  loc_00660010: push eax
  loc_00660011: call [00401080h] ; __vbaHresultCheckObj
  loc_00660017: mov eax, var_74
  loc_0066001A: lea edx, var_60
  loc_0066001D: push edx
  loc_0066001E: push eax
  loc_0066001F: mov ecx, [eax]
  loc_00660021: mov edi, eax
  loc_00660023: call [ecx+00000050h]
  loc_00660026: test eax, eax
  loc_00660028: fnclex
  loc_0066002A: jge 0066003Bh
  loc_0066002C: push 00000050h
  loc_0066002E: push 0042E528h
  loc_00660033: push edi
  loc_00660034: push eax
  loc_00660035: call [00401080h] ; __vbaHresultCheckObj
  loc_0066003B: mov eax, var_60
  loc_0066003E: push eax
  loc_0066003F: push 0043FA7Ch ; "\settings.ini"
  loc_00660044: call [00401070h] ; __vbaStrCat
  loc_0066004A: mov edx, eax
  loc_0066004C: lea ecx, var_6C
  loc_0066004F: call [0040126Ch] ; __vbaStrMove
  loc_00660055: mov edx, 0043FAF0h ; "MachineId"
  loc_0066005A: lea ecx, var_68
  loc_0066005D: call ebx
  loc_0066005F: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_00660064: lea ecx, var_64
  loc_00660067: call ebx
  loc_00660069: mov ecx, [esi]
  loc_0066006B: lea edx, var_70
  loc_0066006E: lea eax, var_6C
  loc_00660071: push edx
  loc_00660072: push eax
  loc_00660073: lea edx, var_68
  loc_00660076: lea eax, var_64
  loc_00660079: push edx
  loc_0066007A: push eax
  loc_0066007B: push esi
  loc_0066007C: call [ecx+000006F8h]
  loc_00660082: test eax, eax
  loc_00660084: jge 00660098h
  loc_00660086: push 000006F8h
  loc_0066008B: push 00434D5Ch
  loc_00660090: push esi
  loc_00660091: push eax
  loc_00660092: call [00401080h] ; __vbaHresultCheckObj
  loc_00660098: mov edx, var_70
  loc_0066009B: lea edi, [esi+00000034h]
  loc_0066009E: mov ecx, edi
  loc_006600A0: call ebx
  loc_006600A2: lea ecx, var_70
  loc_006600A5: lea edx, var_6C
  loc_006600A8: push ecx
  loc_006600A9: lea eax, var_68
  loc_006600AC: push edx
  loc_006600AD: lea ecx, var_64
  loc_006600B0: push eax
  loc_006600B1: lea edx, var_60
  loc_006600B4: push ecx
  loc_006600B5: push edx
  loc_006600B6: push 00000005h
  loc_006600B8: call [00401204h] ; __vbaFreeStrList
  loc_006600BE: add esp, 00000018h
  loc_006600C1: lea ecx, var_74
  loc_006600C4: call [004012A0h] ; __vbaFreeObj
  loc_006600CA: mov eax, [edi]
  loc_006600CC: push eax
  loc_006600CD: push 0042DFECh
  loc_006600D2: call [0040112Ch] ; __vbaStrCmp
  loc_006600D8: test eax, eax
  loc_006600DA: jnz 006600E5h
  loc_006600DC: mov ecx, [esi]
  loc_006600DE: push esi
  loc_006600DF: call [ecx+00000718h]
  loc_006600E5: mov edx, [esi]
  loc_006600E7: push esi
  loc_006600E8: call [edx+00000330h]
  loc_006600EE: push eax
  loc_006600EF: lea eax, var_74
  loc_006600F2: push eax
  loc_006600F3: call [004010B8h] ; __vbaObjSet
  loc_006600F9: mov edx, [edi]
  loc_006600FB: mov ebx, eax
  loc_006600FD: push edx
  loc_006600FE: push ebx
  loc_006600FF: mov ecx, [ebx]
  loc_00660101: call [ecx+000000A4h]
  loc_00660107: test eax, eax
  loc_00660109: fnclex
  loc_0066010B: jge 0066011Fh
  loc_0066010D: push 000000A4h
  loc_00660112: push 0042DFCCh
  loc_00660117: push ebx
  loc_00660118: push eax
  loc_00660119: call [00401080h] ; __vbaHresultCheckObj
  loc_0066011F: mov ebx, [004012A0h] ; __vbaFreeObj
  loc_00660125: lea ecx, var_74
  loc_00660128: call ebx
  loc_0066012A: mov eax, [esi]
  loc_0066012C: push esi
  loc_0066012D: call [eax+0000031Ch]
  loc_00660133: lea ecx, var_74
  loc_00660136: push eax
  loc_00660137: push ecx
  loc_00660138: call [004010B8h] ; __vbaObjSet
  loc_0066013E: mov edi, eax
  loc_00660140: push 0043FC18h ; "Paste Machine ID here"
  loc_00660145: push edi
  loc_00660146: mov edx, [edi]
  loc_00660148: call [edx+000000A4h]
  loc_0066014E: test eax, eax
  loc_00660150: fnclex
  loc_00660152: jge 00660166h
  loc_00660154: push 000000A4h
  loc_00660159: push 0042DFCCh
  loc_0066015E: push edi
  loc_0066015F: push eax
  loc_00660160: call [00401080h] ; __vbaHresultCheckObj
  loc_00660166: lea ecx, var_74
  loc_00660169: call ebx
  loc_0066016B: mov eax, [esi]
  loc_0066016D: push esi
  loc_0066016E: call [eax+00000304h]
  loc_00660174: lea ecx, var_74
  loc_00660177: push eax
  loc_00660178: push ecx
  loc_00660179: call [004010B8h] ; __vbaObjSet
  loc_0066017F: mov edi, eax
  loc_00660181: push 0043FC48h ; "Paste Reg Key here"
  loc_00660186: push edi
  loc_00660187: mov edx, [edi]
  loc_00660189: call [edx+000000A4h]
  loc_0066018F: test eax, eax
  loc_00660191: fnclex
  loc_00660193: jge 006601A7h
  loc_00660195: push 000000A4h
  loc_0066019A: push 0042DFCCh
  loc_0066019F: push edi
  loc_006601A0: push eax
  loc_006601A1: call [00401080h] ; __vbaHresultCheckObj
  loc_006601A7: lea ecx, var_74
  loc_006601AA: call ebx
  loc_006601AC: mov eax, [esi]
  loc_006601AE: push esi
  loc_006601AF: call [eax+0000031Ch]
  loc_006601B5: lea ecx, var_74
  loc_006601B8: push eax
  loc_006601B9: push ecx
  loc_006601BA: call [004010B8h] ; __vbaObjSet
  loc_006601C0: mov edi, eax
  loc_006601C2: push 00000000h
  loc_006601C4: push edi
  loc_006601C5: mov edx, [edi]
  loc_006601C7: call [edx+00000094h]
  loc_006601CD: test eax, eax
  loc_006601CF: fnclex
  loc_006601D1: jge 006601E5h
  loc_006601D3: push 00000094h
  loc_006601D8: push 0042DFCCh
  loc_006601DD: push edi
  loc_006601DE: push eax
  loc_006601DF: call [00401080h] ; __vbaHresultCheckObj
  loc_006601E5: lea ecx, var_74
  loc_006601E8: call ebx
  loc_006601EA: mov eax, [esi]
  loc_006601EC: push esi
  loc_006601ED: call [eax+00000318h]
  loc_006601F3: lea ecx, var_74
  loc_006601F6: push eax
  loc_006601F7: push ecx
  loc_006601F8: call [004010B8h] ; __vbaObjSet
  loc_006601FE: mov edi, eax
  loc_00660200: push 00000000h
  loc_00660202: push edi
  loc_00660203: mov edx, [edi]
  loc_00660205: call [edx+0000008Ch]
  loc_0066020B: test eax, eax
  loc_0066020D: fnclex
  loc_0066020F: jge 00660223h
  loc_00660211: push 0000008Ch
  loc_00660216: push 0043FC70h
  loc_0066021B: push edi
  loc_0066021C: push eax
  loc_0066021D: call [00401080h] ; __vbaHresultCheckObj
  loc_00660223: lea ecx, var_74
  loc_00660226: call ebx
  loc_00660228: mov eax, [esi]
  loc_0066022A: push esi
  loc_0066022B: call [eax+0000071Ch]
  loc_00660231: cmp [esi+0000003Ah], 0000h
  loc_00660236: jnz 006602DBh
  loc_0066023C: mov ecx, 0000000Ah
  loc_00660241: mov eax, 80020004h
  loc_00660246: mov var_B4, ecx
  loc_0066024C: mov var_A4, ecx
  loc_00660252: mov var_94, ecx
  loc_00660258: lea edx, var_D4
  loc_0066025E: lea ecx, var_84
  loc_00660264: mov var_AC, eax
  loc_0066026A: mov var_9C, eax
  loc_00660270: mov var_8C, eax
  loc_00660276: mov var_CC, 0043FC84h ; "Anda sedang menjalankan aplikasi dengan UNREGISTERED mode, bebrapa fungsi akan dinonaktifkan"
  loc_00660280: mov var_D4, 00000008h
  loc_0066028A: call [00401238h] ; __vbaVarDup
  loc_00660290: lea ecx, var_B4
  loc_00660296: lea edx, var_A4
  loc_0066029C: push ecx
  loc_0066029D: lea eax, var_94
  loc_006602A3: push edx
  loc_006602A4: push eax
  loc_006602A5: lea ecx, var_84
  loc_006602AB: push 00000040h
  loc_006602AD: push ecx
  loc_006602AE: call [004010B4h] ; rtcMsgBox
  loc_006602B4: lea edx, var_B4
  loc_006602BA: lea eax, var_A4
  loc_006602C0: push edx
  loc_006602C1: lea ecx, var_94
  loc_006602C7: push eax
  loc_006602C8: lea edx, var_84
  loc_006602CE: push ecx
  loc_006602CF: push edx
  loc_006602D0: push 00000004h
  loc_006602D2: call [0040103Ch] ; __vbaFreeVarList
  loc_006602D8: add esp, 00000014h
  loc_006602DB: mov var_4, 00000000h
  loc_006602E2: fwait
  loc_006602E3: push 006603BDh
  loc_006602E8: jmp 00660341h
  loc_006602EA: lea eax, var_70
  loc_006602ED: lea ecx, var_6C
  loc_006602F0: push eax
  loc_006602F1: lea edx, var_68
  loc_006602F4: push ecx
  loc_006602F5: lea eax, var_64
  loc_006602F8: push edx
  loc_006602F9: lea ecx, var_60
  loc_006602FC: push eax
  loc_006602FD: push ecx
  loc_006602FE: push 00000005h
  loc_00660300: call [00401204h] ; __vbaFreeStrList
  loc_00660306: add esp, 00000018h
  loc_00660309: lea ecx, var_74
  loc_0066030C: call [004012A0h] ; __vbaFreeObj
  loc_00660312: lea edx, var_C4
  loc_00660318: lea eax, var_B4
  loc_0066031E: push edx
  loc_0066031F: lea ecx, var_A4
  loc_00660325: push eax
  loc_00660326: lea edx, var_94
  loc_0066032C: push ecx
  loc_0066032D: lea eax, var_84
  loc_00660333: push edx
  loc_00660334: push eax
  loc_00660335: push 00000005h
  loc_00660337: call [0040103Ch] ; __vbaFreeVarList
  loc_0066033D: add esp, 00000018h
  loc_00660340: ret
  loc_00660341: lea ecx, var_1B4
  loc_00660347: lea edx, var_1A4
  loc_0066034D: push ecx
  loc_0066034E: lea eax, var_194
  loc_00660354: push edx
  loc_00660355: lea ecx, var_184
  loc_0066035B: push eax
  loc_0066035C: lea edx, var_174
  loc_00660362: push ecx
  loc_00660363: lea eax, var_164
  loc_00660369: push edx
  loc_0066036A: lea ecx, var_154
  loc_00660370: push eax
  loc_00660371: lea edx, var_144
  loc_00660377: push ecx
  loc_00660378: push edx
  loc_00660379: push 00000008h
  loc_0066037B: call [0040103Ch] ; __vbaFreeVarList
  loc_00660381: mov esi, [00401020h] ; __vbaFreeVar
  loc_00660387: add esp, 00000024h
  loc_0066038A: lea ecx, var_24
  loc_0066038D: call __vbaFreeVar
  loc_0066038F: lea ecx, var_34
  loc_00660392: call __vbaFreeVar
  loc_00660394: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0066039A: lea ecx, var_38
  loc_0066039D: call __vbaFreeStr
  loc_0066039F: lea ecx, var_118
  loc_006603A5: lea eax, var_50
  loc_006603A8: push ecx
  loc_006603A9: push 00000000h
  loc_006603AB: mov var_118, eax
  loc_006603B1: call [00401090h] ; __vbaAryDestruct
  loc_006603B7: lea ecx, var_58
  loc_006603BA: call __vbaFreeStr
  loc_006603BC: ret
  loc_006603BD: mov eax, Me
  loc_006603C0: push eax
  loc_006603C1: mov edx, [eax]
  loc_006603C3: call [edx+00000008h]
  loc_006603C6: mov eax, var_4
  loc_006603C9: mov ecx, var_14
  loc_006603CC: pop edi
  loc_006603CD: pop esi
  loc_006603CE: mov fs:[00000000h], ecx
  loc_006603D5: pop ebx
  loc_006603D6: mov esp, ebp
  loc_006603D8: pop ebp
  loc_006603D9: retn 0004h
End Sub

Private Sub cmdRegister_Click() '65F240
  loc_0065F240: push ebp
  loc_0065F241: mov ebp, esp
  loc_0065F243: sub esp, 0000000Ch
  loc_0065F246: push 00405696h ; __vbaExceptHandler
  loc_0065F24B: mov eax, fs:[00000000h]
  loc_0065F251: push eax
  loc_0065F252: mov fs:[00000000h], esp
  loc_0065F259: sub esp, 000000FCh
  loc_0065F25F: push ebx
  loc_0065F260: push esi
  loc_0065F261: push edi
  loc_0065F262: mov var_C, esp
  loc_0065F265: mov var_8, 004054A0h
  loc_0065F26C: mov esi, Me
  loc_0065F26F: mov eax, esi
  loc_0065F271: and eax, 00000001h
  loc_0065F274: mov var_4, eax
  loc_0065F277: and esi, FFFFFFFEh
  loc_0065F27A: push esi
  loc_0065F27B: mov Me, esi
  loc_0065F27E: mov ecx, [esi]
  loc_0065F280: call [ecx+00000004h]
  loc_0065F283: mov edi, [00401130h] ; __vbaAryConstruct2
  loc_0065F289: push 00000002h
  loc_0065F28B: lea edx, var_34
  loc_0065F28E: xor ebx, ebx
  loc_0065F290: push 00438BD4h
  loc_0065F295: push edx
  loc_0065F296: mov var_1C, ebx
  loc_0065F299: mov var_58, ebx
  loc_0065F29C: mov var_5C, ebx
  loc_0065F29F: mov var_60, ebx
  loc_0065F2A2: mov var_64, ebx
  loc_0065F2A5: mov var_68, ebx
  loc_0065F2A8: mov var_6C, ebx
  loc_0065F2AB: mov var_7C, ebx
  loc_0065F2AE: mov var_8C, ebx
  loc_0065F2B4: mov var_9C, ebx
  loc_0065F2BA: mov var_AC, ebx
  loc_0065F2C0: mov var_BC, ebx
  loc_0065F2C6: mov var_F0, ebx
  loc_0065F2CC: mov var_F4, ebx
  loc_0065F2D2: call edi
  loc_0065F2D4: push 00000008h
  loc_0065F2D6: lea eax, var_50
  loc_0065F2D9: push 0043FA14h
  loc_0065F2DE: push eax
  loc_0065F2DF: call edi
  loc_0065F2E1: cmp [esi+0000003Ah], bx
  loc_0065F2E5: jnz 0065F67Dh
  loc_0065F2EB: mov ecx, [esi]
  loc_0065F2ED: push esi
  loc_0065F2EE: call [ecx+00000304h]
  loc_0065F2F4: mov ebx, [004010B8h] ; __vbaObjSet
  loc_0065F2FA: lea edx, var_6C
  loc_0065F2FD: push eax
  loc_0065F2FE: push edx
  loc_0065F2FF: call ebx
  loc_0065F301: mov edi, eax
  loc_0065F303: lea ecx, var_58
  loc_0065F306: push ecx
  loc_0065F307: push edi
  loc_0065F308: mov eax, [edi]
  loc_0065F30A: call [eax+000000A0h]
  loc_0065F310: test eax, eax
  loc_0065F312: fnclex
  loc_0065F314: jge 0065F328h
  loc_0065F316: push 000000A0h
  loc_0065F31B: push 0042DFCCh
  loc_0065F320: push edi
  loc_0065F321: push eax
  loc_0065F322: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F328: mov edx, var_58
  loc_0065F32B: push 00000000h
  loc_0065F32D: push FFFFFFFFh
  loc_0065F32F: push 00000001h
  loc_0065F331: push 0042DFECh
  loc_0065F336: push 0042F350h ; "-"
  loc_0065F33B: push edx
  loc_0065F33C: call [00401194h] ; rtcReplace
  loc_0065F342: mov edx, eax
  loc_0065F344: lea ecx, var_1C
  loc_0065F347: call [0040126Ch] ; __vbaStrMove
  loc_0065F34D: lea ecx, var_58
  loc_0065F350: call [0040129Ch] ; __vbaFreeStr
  loc_0065F356: lea ecx, var_6C
  loc_0065F359: call [004012A0h] ; __vbaFreeObj
  loc_0065F35F: mov eax, [esi]
  loc_0065F361: push esi
  loc_0065F362: call [eax+00000308h]
  loc_0065F368: lea ecx, var_6C
  loc_0065F36B: push eax
  loc_0065F36C: push ecx
  loc_0065F36D: call ebx
  loc_0065F36F: mov edi, eax
  loc_0065F371: push 0043FA5Ch ; "Processing..."
  loc_0065F376: push edi
  loc_0065F377: mov edx, [edi]
  loc_0065F379: call [edx+000000A4h]
  loc_0065F37F: test eax, eax
  loc_0065F381: fnclex
  loc_0065F383: jge 0065F397h
  loc_0065F385: push 000000A4h
  loc_0065F38A: push 0042DFCCh
  loc_0065F38F: push edi
  loc_0065F390: push eax
  loc_0065F391: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F397: lea ecx, var_6C
  loc_0065F39A: call [004012A0h] ; __vbaFreeObj
  loc_0065F3A0: mov eax, [esi]
  loc_0065F3A2: push esi
  loc_0065F3A3: call [eax+00000308h]
  loc_0065F3A9: lea ecx, var_6C
  loc_0065F3AC: push eax
  loc_0065F3AD: push ecx
  loc_0065F3AE: call ebx
  loc_0065F3B0: mov edi, eax
  loc_0065F3B2: push edi
  loc_0065F3B3: mov edx, [edi]
  loc_0065F3B5: call [edx+0000021Ch]
  loc_0065F3BB: test eax, eax
  loc_0065F3BD: fnclex
  loc_0065F3BF: jge 0065F3D7h
  loc_0065F3C1: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0065F3C7: push 0000021Ch
  loc_0065F3CC: push 0042DFCCh
  loc_0065F3D1: push edi
  loc_0065F3D2: push eax
  loc_0065F3D3: call ebx
  loc_0065F3D5: jmp 0065F3DDh
  loc_0065F3D7: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0065F3DD: lea ecx, var_6C
  loc_0065F3E0: call [004012A0h] ; __vbaFreeObj
  loc_0065F3E6: mov eax, [0066A318h]
  loc_0065F3EB: test eax, eax
  loc_0065F3ED: jnz 0065F3FFh
  loc_0065F3EF: push 0066A318h
  loc_0065F3F4: push 0042E390h
  loc_0065F3F9: call [004011E8h] ; __vbaNew2
  loc_0065F3FF: mov edi, [0066A318h]
  loc_0065F405: lea ecx, var_6C
  loc_0065F408: push ecx
  loc_0065F409: push edi
  loc_0065F40A: mov eax, [edi]
  loc_0065F40C: call [eax+00000014h]
  loc_0065F40F: test eax, eax
  loc_0065F411: fnclex
  loc_0065F413: jge 0065F420h
  loc_0065F415: push 00000014h
  loc_0065F417: push 0042E380h
  loc_0065F41C: push edi
  loc_0065F41D: push eax
  loc_0065F41E: call ebx
  loc_0065F420: mov eax, var_6C
  loc_0065F423: lea ecx, var_58
  loc_0065F426: push ecx
  loc_0065F427: push eax
  loc_0065F428: mov edx, [eax]
  loc_0065F42A: mov edi, eax
  loc_0065F42C: call [edx+00000050h]
  loc_0065F42F: test eax, eax
  loc_0065F431: fnclex
  loc_0065F433: jge 0065F440h
  loc_0065F435: push 00000050h
  loc_0065F437: push 0042E528h
  loc_0065F43C: push edi
  loc_0065F43D: push eax
  loc_0065F43E: call ebx
  loc_0065F440: mov edx, var_58
  loc_0065F443: push edx
  loc_0065F444: push 0043FA7Ch ; "\settings.ini"
  loc_0065F449: call [00401070h] ; __vbaStrCat
  loc_0065F44F: mov edx, eax
  loc_0065F451: lea ecx, var_64
  loc_0065F454: call [0040126Ch] ; __vbaStrMove
  loc_0065F45A: mov edi, [004011FCh] ; __vbaStrCopy
  loc_0065F460: mov edx, 0043E0BCh ; "RegistrationKey"
  loc_0065F465: lea ecx, var_60
  loc_0065F468: call edi
  loc_0065F46A: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_0065F46F: lea ecx, var_5C
  loc_0065F472: call edi
  loc_0065F474: mov eax, [esi]
  loc_0065F476: lea ecx, var_7C
  loc_0065F479: push ecx
  loc_0065F47A: lea edx, var_64
  loc_0065F47D: lea ecx, var_1C
  loc_0065F480: push edx
  loc_0065F481: push ecx
  loc_0065F482: lea edx, var_60
  loc_0065F485: lea ecx, var_5C
  loc_0065F488: push edx
  loc_0065F489: push ecx
  loc_0065F48A: push esi
  loc_0065F48B: call [eax+000006FCh]
  loc_0065F491: test eax, eax
  loc_0065F493: jge 0065F4A3h
  loc_0065F495: push 000006FCh
  loc_0065F49A: push 00434D5Ch
  loc_0065F49F: push esi
  loc_0065F4A0: push eax
  loc_0065F4A1: call ebx
  loc_0065F4A3: lea edx, var_64
  loc_0065F4A6: lea eax, var_60
  loc_0065F4A9: push edx
  loc_0065F4AA: lea ecx, var_5C
  loc_0065F4AD: push eax
  loc_0065F4AE: lea edx, var_58
  loc_0065F4B1: push ecx
  loc_0065F4B2: push edx
  loc_0065F4B3: push 00000004h
  loc_0065F4B5: call [00401204h] ; __vbaFreeStrList
  loc_0065F4BB: add esp, 00000014h
  loc_0065F4BE: lea ecx, var_6C
  loc_0065F4C1: call [004012A0h] ; __vbaFreeObj
  loc_0065F4C7: lea ecx, var_7C
  loc_0065F4CA: call [00401020h] ; __vbaFreeVar
  loc_0065F4D0: mov eax, [esi]
  loc_0065F4D2: push esi
  loc_0065F4D3: call [eax+0000071Ch]
  loc_0065F4D9: cmp [esi+0000003Ah], 0000h
  loc_0065F4DE: jnz 0065F89Eh
  loc_0065F4E4: mov ecx, 80020004h
  loc_0065F4E9: mov eax, 0000000Ah
  loc_0065F4EE: mov var_A4, ecx
  loc_0065F4F4: mov var_94, ecx
  loc_0065F4FA: mov var_84, ecx
  loc_0065F500: lea edx, var_BC
  loc_0065F506: lea ecx, var_7C
  loc_0065F509: mov var_AC, eax
  loc_0065F50F: mov var_9C, eax
  loc_0065F515: mov var_8C, eax
  loc_0065F51B: mov var_B4, 0043FA9Ch ; "Anda salah memasukkan kode registrati "
  loc_0065F525: mov var_BC, 00000008h
  loc_0065F52F: call [00401238h] ; __vbaVarDup
  loc_0065F535: lea ecx, var_AC
  loc_0065F53B: lea edx, var_9C
  loc_0065F541: push ecx
  loc_0065F542: lea eax, var_8C
  loc_0065F548: push edx
  loc_0065F549: push eax
  loc_0065F54A: lea ecx, var_7C
  loc_0065F54D: push 00000010h
  loc_0065F54F: push ecx
  loc_0065F550: call [004010B4h] ; rtcMsgBox
  loc_0065F556: lea edx, var_AC
  loc_0065F55C: lea eax, var_9C
  loc_0065F562: push edx
  loc_0065F563: lea ecx, var_8C
  loc_0065F569: push eax
  loc_0065F56A: lea edx, var_7C
  loc_0065F56D: push ecx
  loc_0065F56E: push edx
  loc_0065F56F: push 00000004h
  loc_0065F571: call [0040103Ch] ; __vbaFreeVarList
  loc_0065F577: mov eax, [0066A318h]
  loc_0065F57C: add esp, 00000014h
  loc_0065F57F: test eax, eax
  loc_0065F581: jnz 0065F593h
  loc_0065F583: push 0066A318h
  loc_0065F588: push 0042E390h
  loc_0065F58D: call [004011E8h] ; __vbaNew2
  loc_0065F593: mov ebx, [0066A318h]
  loc_0065F599: lea ecx, var_6C
  loc_0065F59C: push ecx
  loc_0065F59D: push ebx
  loc_0065F59E: mov eax, [ebx]
  loc_0065F5A0: call [eax+00000014h]
  loc_0065F5A3: test eax, eax
  loc_0065F5A5: fnclex
  loc_0065F5A7: jge 0065F5B8h
  loc_0065F5A9: push 00000014h
  loc_0065F5AB: push 0042E380h
  loc_0065F5B0: push ebx
  loc_0065F5B1: push eax
  loc_0065F5B2: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F5B8: mov eax, var_6C
  loc_0065F5BB: lea ecx, var_58
  loc_0065F5BE: push ecx
  loc_0065F5BF: push eax
  loc_0065F5C0: mov edx, [eax]
  loc_0065F5C2: mov ebx, eax
  loc_0065F5C4: call [edx+00000050h]
  loc_0065F5C7: test eax, eax
  loc_0065F5C9: fnclex
  loc_0065F5CB: jge 0065F5DCh
  loc_0065F5CD: push 00000050h
  loc_0065F5CF: push 0042E528h
  loc_0065F5D4: push ebx
  loc_0065F5D5: push eax
  loc_0065F5D6: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F5DC: mov edx, var_58
  loc_0065F5DF: push edx
  loc_0065F5E0: push 0043FA7Ch ; "\settings.ini"
  loc_0065F5E5: call [00401070h] ; __vbaStrCat
  loc_0065F5EB: mov edx, eax
  loc_0065F5ED: lea ecx, var_68
  loc_0065F5F0: call [0040126Ch] ; __vbaStrMove
  loc_0065F5F6: mov edx, 0042DFECh
  loc_0065F5FB: lea ecx, var_64
  loc_0065F5FE: call edi
  loc_0065F600: mov edx, 0043E0BCh ; "RegistrationKey"
  loc_0065F605: lea ecx, var_60
  loc_0065F608: call edi
  loc_0065F60A: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_0065F60F: lea ecx, var_5C
  loc_0065F612: call edi
  loc_0065F614: mov eax, [esi]
  loc_0065F616: lea ecx, var_7C
  loc_0065F619: push ecx
  loc_0065F61A: lea edx, var_68
  loc_0065F61D: lea ecx, var_64
  loc_0065F620: push edx
  loc_0065F621: push ecx
  loc_0065F622: lea edx, var_60
  loc_0065F625: lea ecx, var_5C
  loc_0065F628: push edx
  loc_0065F629: push ecx
  loc_0065F62A: push esi
  loc_0065F62B: call [eax+000006FCh]
  loc_0065F631: test eax, eax
  loc_0065F633: jge 0065F647h
  loc_0065F635: push 000006FCh
  loc_0065F63A: push 00434D5Ch
  loc_0065F63F: push esi
  loc_0065F640: push eax
  loc_0065F641: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F647: lea edx, var_68
  loc_0065F64A: lea eax, var_64
  loc_0065F64D: push edx
  loc_0065F64E: lea ecx, var_60
  loc_0065F651: push eax
  loc_0065F652: lea edx, var_5C
  loc_0065F655: push ecx
  loc_0065F656: lea eax, var_58
  loc_0065F659: push edx
  loc_0065F65A: push eax
  loc_0065F65B: push 00000005h
  loc_0065F65D: call [00401204h] ; __vbaFreeStrList
  loc_0065F663: add esp, 00000018h
  loc_0065F666: lea ecx, var_6C
  loc_0065F669: call [004012A0h] ; __vbaFreeObj
  loc_0065F66F: lea ecx, var_7C
  loc_0065F672: call [00401020h] ; __vbaFreeVar
  loc_0065F678: jmp 0065F89Eh
  loc_0065F67D: cmp [0066A318h], ebx
  loc_0065F683: jnz 0065F695h
  loc_0065F685: push 0066A318h
  loc_0065F68A: push 0042E390h
  loc_0065F68F: call [004011E8h] ; __vbaNew2
  loc_0065F695: mov edi, [0066A318h]
  loc_0065F69B: lea edx, var_6C
  loc_0065F69E: push edx
  loc_0065F69F: push edi
  loc_0065F6A0: mov ecx, [edi]
  loc_0065F6A2: call [ecx+00000014h]
  loc_0065F6A5: cmp eax, ebx
  loc_0065F6A7: fnclex
  loc_0065F6A9: jge 0065F6BAh
  loc_0065F6AB: push 00000014h
  loc_0065F6AD: push 0042E380h
  loc_0065F6B2: push edi
  loc_0065F6B3: push eax
  loc_0065F6B4: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F6BA: mov eax, var_6C
  loc_0065F6BD: lea edx, var_58
  loc_0065F6C0: push edx
  loc_0065F6C1: push eax
  loc_0065F6C2: mov ecx, [eax]
  loc_0065F6C4: mov edi, eax
  loc_0065F6C6: call [ecx+00000050h]
  loc_0065F6C9: cmp eax, ebx
  loc_0065F6CB: fnclex
  loc_0065F6CD: jge 0065F6DEh
  loc_0065F6CF: push 00000050h
  loc_0065F6D1: push 0042E528h
  loc_0065F6D6: push edi
  loc_0065F6D7: push eax
  loc_0065F6D8: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F6DE: mov eax, var_58
  loc_0065F6E1: push eax
  loc_0065F6E2: push 0043FA7Ch ; "\settings.ini"
  loc_0065F6E7: call [00401070h] ; __vbaStrCat
  loc_0065F6ED: mov edx, eax
  loc_0065F6EF: lea ecx, var_68
  loc_0065F6F2: call [0040126Ch] ; __vbaStrMove
  loc_0065F6F8: mov edi, [004011FCh] ; __vbaStrCopy
  loc_0065F6FE: mov edx, 0042DFECh
  loc_0065F703: lea ecx, var_64
  loc_0065F706: call edi
  loc_0065F708: mov edx, 0043FAF0h ; "MachineId"
  loc_0065F70D: lea ecx, var_60
  loc_0065F710: call edi
  loc_0065F712: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_0065F717: lea ecx, var_5C
  loc_0065F71A: call edi
  loc_0065F71C: mov ecx, [esi]
  loc_0065F71E: lea edx, var_7C
  loc_0065F721: push edx
  loc_0065F722: lea eax, var_68
  loc_0065F725: lea edx, var_64
  loc_0065F728: push eax
  loc_0065F729: push edx
  loc_0065F72A: lea eax, var_60
  loc_0065F72D: lea edx, var_5C
  loc_0065F730: push eax
  loc_0065F731: push edx
  loc_0065F732: push esi
  loc_0065F733: call [ecx+000006FCh]
  loc_0065F739: cmp eax, ebx
  loc_0065F73B: jge 0065F74Fh
  loc_0065F73D: push 000006FCh
  loc_0065F742: push 00434D5Ch
  loc_0065F747: push esi
  loc_0065F748: push eax
  loc_0065F749: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F74F: lea eax, var_68
  loc_0065F752: lea ecx, var_64
  loc_0065F755: push eax
  loc_0065F756: lea edx, var_60
  loc_0065F759: push ecx
  loc_0065F75A: lea eax, var_5C
  loc_0065F75D: push edx
  loc_0065F75E: lea ecx, var_58
  loc_0065F761: push eax
  loc_0065F762: push ecx
  loc_0065F763: push 00000005h
  loc_0065F765: call [00401204h] ; __vbaFreeStrList
  loc_0065F76B: add esp, 00000018h
  loc_0065F76E: lea ecx, var_6C
  loc_0065F771: call [004012A0h] ; __vbaFreeObj
  loc_0065F777: lea ecx, var_7C
  loc_0065F77A: call [00401020h] ; __vbaFreeVar
  loc_0065F780: cmp [0066A318h], ebx
  loc_0065F786: jnz 0065F798h
  loc_0065F788: push 0066A318h
  loc_0065F78D: push 0042E390h
  loc_0065F792: call [004011E8h] ; __vbaNew2
  loc_0065F798: mov ebx, [0066A318h]
  loc_0065F79E: lea eax, var_6C
  loc_0065F7A1: push eax
  loc_0065F7A2: push ebx
  loc_0065F7A3: mov edx, [ebx]
  loc_0065F7A5: call [edx+00000014h]
  loc_0065F7A8: test eax, eax
  loc_0065F7AA: fnclex
  loc_0065F7AC: jge 0065F7BDh
  loc_0065F7AE: push 00000014h
  loc_0065F7B0: push 0042E380h
  loc_0065F7B5: push ebx
  loc_0065F7B6: push eax
  loc_0065F7B7: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F7BD: mov eax, var_6C
  loc_0065F7C0: lea edx, var_58
  loc_0065F7C3: push edx
  loc_0065F7C4: push eax
  loc_0065F7C5: mov ecx, [eax]
  loc_0065F7C7: mov ebx, eax
  loc_0065F7C9: call [ecx+00000050h]
  loc_0065F7CC: test eax, eax
  loc_0065F7CE: fnclex
  loc_0065F7D0: jge 0065F7E5h
  loc_0065F7D2: push 00000050h
  loc_0065F7D4: push 0042E528h
  loc_0065F7D9: push ebx
  loc_0065F7DA: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0065F7E0: push eax
  loc_0065F7E1: call ebx
  loc_0065F7E3: jmp 0065F7EBh
  loc_0065F7E5: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0065F7EB: mov eax, var_58
  loc_0065F7EE: push eax
  loc_0065F7EF: push 0043FA7Ch ; "\settings.ini"
  loc_0065F7F4: call [00401070h] ; __vbaStrCat
  loc_0065F7FA: mov edx, eax
  loc_0065F7FC: lea ecx, var_68
  loc_0065F7FF: call [0040126Ch] ; __vbaStrMove
  loc_0065F805: mov edx, 0042DFECh
  loc_0065F80A: lea ecx, var_64
  loc_0065F80D: call edi
  loc_0065F80F: mov edx, 0043E0BCh ; "RegistrationKey"
  loc_0065F814: lea ecx, var_60
  loc_0065F817: call edi
  loc_0065F819: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_0065F81E: lea ecx, var_5C
  loc_0065F821: call edi
  loc_0065F823: mov ecx, [esi]
  loc_0065F825: lea edx, var_7C
  loc_0065F828: push edx
  loc_0065F829: lea eax, var_68
  loc_0065F82C: lea edx, var_64
  loc_0065F82F: push eax
  loc_0065F830: push edx
  loc_0065F831: lea eax, var_60
  loc_0065F834: lea edx, var_5C
  loc_0065F837: push eax
  loc_0065F838: push edx
  loc_0065F839: push esi
  loc_0065F83A: call [ecx+000006FCh]
  loc_0065F840: test eax, eax
  loc_0065F842: jge 0065F852h
  loc_0065F844: push 000006FCh
  loc_0065F849: push 00434D5Ch
  loc_0065F84E: push esi
  loc_0065F84F: push eax
  loc_0065F850: call ebx
  loc_0065F852: lea eax, var_68
  loc_0065F855: lea ecx, var_64
  loc_0065F858: push eax
  loc_0065F859: lea edx, var_60
  loc_0065F85C: push ecx
  loc_0065F85D: lea eax, var_5C
  loc_0065F860: push edx
  loc_0065F861: lea ecx, var_58
  loc_0065F864: push eax
  loc_0065F865: push ecx
  loc_0065F866: push 00000005h
  loc_0065F868: call [00401204h] ; __vbaFreeStrList
  loc_0065F86E: add esp, 00000018h
  loc_0065F871: lea ecx, var_6C
  loc_0065F874: call [004012A0h] ; __vbaFreeObj
  loc_0065F87A: lea ecx, var_7C
  loc_0065F87D: call [00401020h] ; __vbaFreeVar
  loc_0065F883: mov edx, [esi]
  loc_0065F885: push esi
  loc_0065F886: call [edx+00000714h]
  loc_0065F88C: test eax, eax
  loc_0065F88E: jge 0065F89Eh
  loc_0065F890: push 00000714h
  loc_0065F895: push 00434D5Ch
  loc_0065F89A: push esi
  loc_0065F89B: push eax
  loc_0065F89C: call ebx
  loc_0065F89E: mov var_4, 00000000h
  loc_0065F8A5: push 0065F931h
  loc_0065F8AA: jmp 0065F8F9h
  loc_0065F8AC: lea eax, var_68
  loc_0065F8AF: lea ecx, var_64
  loc_0065F8B2: push eax
  loc_0065F8B3: lea edx, var_60
  loc_0065F8B6: push ecx
  loc_0065F8B7: lea eax, var_5C
  loc_0065F8BA: push edx
  loc_0065F8BB: lea ecx, var_58
  loc_0065F8BE: push eax
  loc_0065F8BF: push ecx
  loc_0065F8C0: push 00000005h
  loc_0065F8C2: call [00401204h] ; __vbaFreeStrList
  loc_0065F8C8: add esp, 00000018h
  loc_0065F8CB: lea ecx, var_6C
  loc_0065F8CE: call [004012A0h] ; __vbaFreeObj
  loc_0065F8D4: lea edx, var_AC
  loc_0065F8DA: lea eax, var_9C
  loc_0065F8E0: push edx
  loc_0065F8E1: lea ecx, var_8C
  loc_0065F8E7: push eax
  loc_0065F8E8: lea edx, var_7C
  loc_0065F8EB: push ecx
  loc_0065F8EC: push edx
  loc_0065F8ED: push 00000004h
  loc_0065F8EF: call [0040103Ch] ; __vbaFreeVarList
  loc_0065F8F5: add esp, 00000014h
  loc_0065F8F8: ret
  loc_0065F8F9: lea ecx, var_1C
  loc_0065F8FC: call [0040129Ch] ; __vbaFreeStr
  loc_0065F902: mov esi, [00401090h] ; __vbaAryDestruct
  loc_0065F908: lea ecx, var_F0
  loc_0065F90E: lea eax, var_34
  loc_0065F911: push ecx
  loc_0065F912: push 00000000h
  loc_0065F914: mov var_F0, eax
  loc_0065F91A: call __vbaAryDestruct
  loc_0065F91C: lea eax, var_F4
  loc_0065F922: lea edx, var_50
  loc_0065F925: push eax
  loc_0065F926: push 00000000h
  loc_0065F928: mov var_F4, edx
  loc_0065F92E: call __vbaAryDestruct
  loc_0065F930: ret
  loc_0065F931: mov eax, Me
  loc_0065F934: push eax
  loc_0065F935: mov ecx, [eax]
  loc_0065F937: call [ecx+00000008h]
  loc_0065F93A: mov eax, var_4
  loc_0065F93D: mov ecx, var_14
  loc_0065F940: pop edi
  loc_0065F941: pop esi
  loc_0065F942: mov fs:[00000000h], ecx
  loc_0065F949: pop ebx
  loc_0065F94A: mov esp, ebp
  loc_0065F94C: pop ebp
  loc_0065F94D: retn 0004h
End Sub

Private Sub cmdCopyId_Click() '65EEE0
  loc_0065EEE0: push ebp
  loc_0065EEE1: mov ebp, esp
  loc_0065EEE3: sub esp, 0000000Ch
  loc_0065EEE6: push 00405696h ; __vbaExceptHandler
  loc_0065EEEB: mov eax, fs:[00000000h]
  loc_0065EEF1: push eax
  loc_0065EEF2: mov fs:[00000000h], esp
  loc_0065EEF9: sub esp, 0000003Ch
  loc_0065EEFC: push ebx
  loc_0065EEFD: push esi
  loc_0065EEFE: push edi
  loc_0065EEFF: mov var_C, esp
  loc_0065EF02: mov var_8, 00405480h
  loc_0065EF09: mov edi, Me
  loc_0065EF0C: mov eax, edi
  loc_0065EF0E: and eax, 00000001h
  loc_0065EF11: mov var_4, eax
  loc_0065EF14: and edi, FFFFFFFEh
  loc_0065EF17: push edi
  loc_0065EF18: mov Me, edi
  loc_0065EF1B: mov ecx, [edi]
  loc_0065EF1D: call [ecx+00000004h]
  loc_0065EF20: mov eax, [0066A318h]
  loc_0065EF25: xor ebx, ebx
  loc_0065EF27: cmp eax, ebx
  loc_0065EF29: mov var_18, ebx
  loc_0065EF2C: mov var_1C, ebx
  loc_0065EF2F: mov var_20, ebx
  loc_0065EF32: jnz 0065EF44h
  loc_0065EF34: push 0066A318h
  loc_0065EF39: push 0042E390h
  loc_0065EF3E: call [004011E8h] ; __vbaNew2
  loc_0065EF44: mov esi, [0066A318h]
  loc_0065EF4A: lea eax, var_20
  loc_0065EF4D: push eax
  loc_0065EF4E: push esi
  loc_0065EF4F: mov edx, [esi]
  loc_0065EF51: call [edx+0000001Ch]
  loc_0065EF54: cmp eax, ebx
  loc_0065EF56: fnclex
  loc_0065EF58: jge 0065EF69h
  loc_0065EF5A: push 0000001Ch
  loc_0065EF5C: push 0042E380h
  loc_0065EF61: push esi
  loc_0065EF62: push eax
  loc_0065EF63: call [00401080h] ; __vbaHresultCheckObj
  loc_0065EF69: mov ecx, [edi]
  loc_0065EF6B: mov ebx, var_20
  loc_0065EF6E: push edi
  loc_0065EF6F: call [ecx+00000330h]
  loc_0065EF75: lea edx, var_1C
  loc_0065EF78: push eax
  loc_0065EF79: push edx
  loc_0065EF7A: call [004010B8h] ; __vbaObjSet
  loc_0065EF80: mov esi, eax
  loc_0065EF82: lea ecx, var_18
  loc_0065EF85: push ecx
  loc_0065EF86: push esi
  loc_0065EF87: mov eax, [esi]
  loc_0065EF89: call [eax+000000A0h]
  loc_0065EF8F: test eax, eax
  loc_0065EF91: fnclex
  loc_0065EF93: jge 0065EFA7h
  loc_0065EF95: push 000000A0h
  loc_0065EF9A: push 0042DFCCh
  loc_0065EF9F: push esi
  loc_0065EFA0: push eax
  loc_0065EFA1: call [00401080h] ; __vbaHresultCheckObj
  loc_0065EFA7: sub esp, 00000010h
  loc_0065EFAA: mov eax, 0000000Ah
  loc_0065EFAF: mov ecx, esp
  loc_0065EFB1: mov edx, [ebx]
  loc_0065EFB3: mov [ecx], eax
  loc_0065EFB5: mov eax, var_2C
  loc_0065EFB8: mov [ecx+00000004h], eax
  loc_0065EFBB: mov eax, 80020004h
  loc_0065EFC0: mov [ecx+00000008h], eax
  loc_0065EFC3: mov eax, var_24
  loc_0065EFC6: mov [ecx+0000000Ch], eax
  loc_0065EFC9: mov ecx, var_18
  loc_0065EFCC: push ecx
  loc_0065EFCD: push ebx
  loc_0065EFCE: call [edx+00000060h]
  loc_0065EFD1: test eax, eax
  loc_0065EFD3: fnclex
  loc_0065EFD5: jge 0065EFE6h
  loc_0065EFD7: push 00000060h
  loc_0065EFD9: push 0043FA48h
  loc_0065EFDE: push ebx
  loc_0065EFDF: push eax
  loc_0065EFE0: call [00401080h] ; __vbaHresultCheckObj
  loc_0065EFE6: lea ecx, var_18
  loc_0065EFE9: call [0040129Ch] ; __vbaFreeStr
  loc_0065EFEF: lea edx, var_20
  loc_0065EFF2: lea eax, var_1C
  loc_0065EFF5: push edx
  loc_0065EFF6: push eax
  loc_0065EFF7: push 00000002h
  loc_0065EFF9: call [0040104Ch] ; __vbaFreeObjList
  loc_0065EFFF: mov ecx, [edi]
  loc_0065F001: add esp, 0000000Ch
  loc_0065F004: push edi
  loc_0065F005: call [ecx+0000031Ch]
  loc_0065F00B: lea edx, var_1C
  loc_0065F00E: push eax
  loc_0065F00F: push edx
  loc_0065F010: call [004010B8h] ; __vbaObjSet
  loc_0065F016: mov esi, eax
  loc_0065F018: push FFFFFFFFh
  loc_0065F01A: push esi
  loc_0065F01B: mov eax, [esi]
  loc_0065F01D: call [eax+00000094h]
  loc_0065F023: test eax, eax
  loc_0065F025: fnclex
  loc_0065F027: jge 0065F03Bh
  loc_0065F029: push 00000094h
  loc_0065F02E: push 0042DFCCh
  loc_0065F033: push esi
  loc_0065F034: push eax
  loc_0065F035: call [00401080h] ; __vbaHresultCheckObj
  loc_0065F03B: lea ecx, var_1C
  loc_0065F03E: call [004012A0h] ; __vbaFreeObj
  loc_0065F044: mov var_4, 00000000h
  loc_0065F04B: push 0065F070h
  loc_0065F050: jmp 0065F06Fh
  loc_0065F052: lea ecx, var_18
  loc_0065F055: call [0040129Ch] ; __vbaFreeStr
  loc_0065F05B: lea ecx, var_20
  loc_0065F05E: lea edx, var_1C
  loc_0065F061: push ecx
  loc_0065F062: push edx
  loc_0065F063: push 00000002h
  loc_0065F065: call [0040104Ch] ; __vbaFreeObjList
  loc_0065F06B: add esp, 0000000Ch
  loc_0065F06E: ret
  loc_0065F06F: ret
  loc_0065F070: mov eax, Me
  loc_0065F073: push eax
  loc_0065F074: mov ecx, [eax]
  loc_0065F076: call [ecx+00000008h]
  loc_0065F079: mov eax, var_4
  loc_0065F07C: mov ecx, var_14
  loc_0065F07F: pop edi
  loc_0065F080: pop esi
  loc_0065F081: mov fs:[00000000h], ecx
  loc_0065F088: pop ebx
  loc_0065F089: mov esp, ebp
  loc_0065F08B: pop ebp
  loc_0065F08C: retn 0004h
End Sub

Public Function ReadINI(strsection, strkey, strfullpath) '662490
  loc_00662490: push ebp
  loc_00662491: mov ebp, esp
  loc_00662493: sub esp, 0000000Ch
  loc_00662496: push 00405696h ; __vbaExceptHandler
  loc_0066249B: mov eax, fs:[00000000h]
  loc_006624A1: push eax
  loc_006624A2: mov fs:[00000000h], esp
  loc_006624A9: sub esp, 00000044h
  loc_006624AC: push ebx
  loc_006624AD: push esi
  loc_006624AE: push edi
  loc_006624AF: mov var_C, esp
  loc_006624B2: mov var_8, 00405538h
  loc_006624B9: xor esi, esi
  loc_006624BB: mov var_4, esi
  loc_006624BE: mov eax, Me
  loc_006624C1: push eax
  loc_006624C2: mov ecx, [eax]
  loc_006624C4: call [ecx+00000004h]
  loc_006624C7: mov edx, arg_18
  loc_006624CA: push esi
  loc_006624CB: mov var_18, esi
  loc_006624CE: mov var_1C, esi
  loc_006624D1: mov var_20, esi
  loc_006624D4: mov var_24, esi
  loc_006624D7: mov var_28, esi
  loc_006624DA: mov var_2C, esi
  loc_006624DD: mov var_30, esi
  loc_006624E0: mov var_34, esi
  loc_006624E3: mov var_38, esi
  loc_006624E6: mov var_48, esi
  loc_006624E9: mov [edx], esi
  loc_006624EB: call [004011D8h] ; rtcBstrFromAnsi
  loc_006624F1: mov var_40, eax
  loc_006624F4: lea eax, var_48
  loc_006624F7: push eax
  loc_006624F8: push 00000020h
  loc_006624FA: mov var_48, 00000008h
  loc_00662501: call [00401198h] ; rtcStringBstr
  loc_00662507: mov edi, [0040126Ch] ; __vbaStrMove
  loc_0066250D: mov edx, eax
  loc_0066250F: lea ecx, var_18
  loc_00662512: call edi
  loc_00662514: lea ecx, var_48
  loc_00662517: call [00401020h] ; __vbaFreeVar
  loc_0066251D: mov ecx, strkey
  loc_00662520: mov edx, [ecx]
  loc_00662522: push edx
  loc_00662523: call [00401058h] ; rtcLowerCaseBstr
  loc_00662529: mov edx, eax
  loc_0066252B: lea ecx, var_38
  loc_0066252E: call edi
  loc_00662530: mov eax, strfullpath
  loc_00662533: mov ebx, var_38
  loc_00662536: lea edx, var_34
  loc_00662539: mov var_38, esi
  loc_0066253C: mov ecx, [eax]
  loc_0066253E: mov esi, [00401230h] ; __vbaStrToAnsi
  loc_00662544: push ecx
  loc_00662545: push edx
  loc_00662546: call __vbaStrToAnsi
  loc_00662548: push eax
  loc_00662549: mov eax, var_18
  loc_0066254C: push eax
  loc_0066254D: call [00401030h] ; __vbaLenBstr
  loc_00662553: mov ecx, var_18
  loc_00662556: push eax
  loc_00662557: lea edx, var_30
  loc_0066255A: push ecx
  loc_0066255B: push edx
  loc_0066255C: call __vbaStrToAnsi
  loc_0066255E: push eax
  loc_0066255F: lea eax, var_2C
  loc_00662562: push 0042DFECh
  loc_00662567: push eax
  loc_00662568: call __vbaStrToAnsi
  loc_0066256A: push eax
  loc_0066256B: mov edx, ebx
  loc_0066256D: lea ecx, var_24
  loc_00662570: call edi
  loc_00662572: lea ecx, var_28
  loc_00662575: push eax
  loc_00662576: push ecx
  loc_00662577: call __vbaStrToAnsi
  loc_00662579: mov ebx, strsection
  loc_0066257C: push eax
  loc_0066257D: lea eax, var_20
  loc_00662580: mov edx, [ebx]
  loc_00662582: push edx
  loc_00662583: push eax
  loc_00662584: call __vbaStrToAnsi
  loc_00662586: push eax
  loc_00662587: GetPrivateProfileString(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v)
  loc_0066258C: mov var_4C, eax
  loc_0066258F: call [0040107Ch] ; __vbaSetSystemError
  loc_00662595: mov ecx, var_20
  loc_00662598: mov esi, [00401190h] ; __vbaStrToUnicode
  loc_0066259E: push ecx
  loc_0066259F: push ebx
  loc_006625A0: call __vbaStrToUnicode
  loc_006625A2: mov edx, var_30
  loc_006625A5: lea eax, var_18
  loc_006625A8: push edx
  loc_006625A9: push eax
  loc_006625AA: call __vbaStrToUnicode
  loc_006625AC: mov ecx, var_34
  loc_006625AF: mov edx, strfullpath
  loc_006625B2: push ecx
  loc_006625B3: push edx
  loc_006625B4: call __vbaStrToUnicode
  loc_006625B6: mov eax, var_4C
  loc_006625B9: mov ecx, var_18
  loc_006625BC: push eax
  loc_006625BD: push ecx
  loc_006625BE: call [0040124Ch] ; rtcLeftCharBstr
  loc_006625C4: mov edx, eax
  loc_006625C6: lea ecx, var_1C
  loc_006625C9: call edi
  loc_006625CB: lea edx, var_38
  loc_006625CE: lea eax, var_34
  loc_006625D1: push edx
  loc_006625D2: lea ecx, var_30
  loc_006625D5: push eax
  loc_006625D6: lea edx, var_2C
  loc_006625D9: push ecx
  loc_006625DA: lea eax, var_28
  loc_006625DD: push edx
  loc_006625DE: lea ecx, var_24
  loc_006625E1: push eax
  loc_006625E2: lea edx, var_20
  loc_006625E5: push ecx
  loc_006625E6: push edx
  loc_006625E7: push 00000007h
  loc_006625E9: call [00401204h] ; __vbaFreeStrList
  loc_006625EF: add esp, 00000020h
  loc_006625F2: push 00662643h
  loc_006625F7: jmp 00662639h
  loc_006625F9: test var_4, 04h
  loc_006625FD: jz 00662608h
  loc_006625FF: lea ecx, var_1C
  loc_00662602: call [0040129Ch] ; __vbaFreeStr
  loc_00662608: lea eax, var_38
  loc_0066260B: lea ecx, var_34
  loc_0066260E: push eax
  loc_0066260F: lea edx, var_30
  loc_00662612: push ecx
  loc_00662613: lea eax, var_2C
  loc_00662616: push edx
  loc_00662617: lea ecx, var_28
  loc_0066261A: push eax
  loc_0066261B: lea edx, var_24
  loc_0066261E: push ecx
  loc_0066261F: lea eax, var_20
  loc_00662622: push edx
  loc_00662623: push eax
  loc_00662624: push 00000007h
  loc_00662626: call [00401204h] ; __vbaFreeStrList
  loc_0066262C: add esp, 00000020h
  loc_0066262F: lea ecx, var_48
  loc_00662632: call [00401020h] ; __vbaFreeVar
  loc_00662638: ret
  loc_00662639: lea ecx, var_18
  loc_0066263C: call [0040129Ch] ; __vbaFreeStr
  loc_00662642: ret
  loc_00662643: mov eax, Me
  loc_00662646: push eax
  loc_00662647: mov ecx, [eax]
  loc_00662649: call [ecx+00000008h]
  loc_0066264C: mov edx, arg_18
  loc_0066264F: mov eax, var_1C
  loc_00662652: mov [edx], eax
  loc_00662654: mov eax, var_4
  loc_00662657: mov ecx, var_14
  loc_0066265A: pop edi
  loc_0066265B: pop esi
  loc_0066265C: mov fs:[00000000h], ecx
  loc_00662663: pop ebx
  loc_00662664: mov esp, ebp
  loc_00662666: pop ebp
  loc_00662667: retn 0014h
End Function

Public Function WriteINI(strsection, strkey, strkeyvalue, strfullpath) '662670
  loc_00662670: push ebp
  loc_00662671: mov ebp, esp
  loc_00662673: sub esp, 0000000Ch
  loc_00662676: push 00405696h ; __vbaExceptHandler
  loc_0066267B: mov eax, fs:[00000000h]
  loc_00662681: push eax
  loc_00662682: mov fs:[00000000h], esp
  loc_00662689: sub esp, 00000034h
  loc_0066268C: push ebx
  loc_0066268D: push esi
  loc_0066268E: push edi
  loc_0066268F: mov var_C, esp
  loc_00662692: mov var_8, 00405548h
  loc_00662699: xor esi, esi
  loc_0066269B: mov var_4, esi
  loc_0066269E: mov eax, Me
  loc_006626A1: push eax
  loc_006626A2: mov ecx, [eax]
  loc_006626A4: call [ecx+00000004h]
  loc_006626A7: mov edx, arg_1C
  loc_006626AA: mov eax, strkey
  loc_006626AD: mov var_24, esi
  loc_006626B0: mov var_28, esi
  loc_006626B3: mov [edx], esi
  loc_006626B5: mov ecx, [eax]
  loc_006626B7: push ecx
  loc_006626B8: mov var_2C, esi
  loc_006626BB: mov var_30, esi
  loc_006626BE: mov var_34, esi
  loc_006626C1: mov var_38, esi
  loc_006626C4: mov var_3C, esi
  loc_006626C7: call [0040111Ch] ; rtcUpperCaseBstr
  loc_006626CD: mov edi, [0040126Ch] ; __vbaStrMove
  loc_006626D3: mov edx, eax
  loc_006626D5: lea ecx, var_3C
  loc_006626D8: call edi
  loc_006626DA: mov edx, strfullpath
  loc_006626DD: mov ebx, var_3C
  loc_006626E0: lea ecx, var_38
  loc_006626E3: mov var_3C, esi
  loc_006626E6: mov eax, [edx]
  loc_006626E8: mov esi, [00401230h] ; __vbaStrToAnsi
  loc_006626EE: push eax
  loc_006626EF: push ecx
  loc_006626F0: call __vbaStrToAnsi
  loc_006626F2: mov edx, strkeyvalue
  loc_006626F5: push eax
  loc_006626F6: lea ecx, var_34
  loc_006626F9: mov eax, [edx]
  loc_006626FB: push eax
  loc_006626FC: push ecx
  loc_006626FD: call __vbaStrToAnsi
  loc_006626FF: push eax
  loc_00662700: mov edx, ebx
  loc_00662702: lea ecx, var_2C
  loc_00662705: call edi
  loc_00662707: lea edx, var_30
  loc_0066270A: push eax
  loc_0066270B: push edx
  loc_0066270C: call __vbaStrToAnsi
  loc_0066270E: mov edi, strsection
  loc_00662711: push eax
  loc_00662712: lea ecx, var_28
  loc_00662715: mov eax, [edi]
  loc_00662717: push eax
  loc_00662718: push ecx
  loc_00662719: call __vbaStrToAnsi
  loc_0066271B: push eax
  loc_0066271C: WritePrivateProfileString(%x1v, %x2v, %x3v, %x4v)
  loc_00662721: call [0040107Ch] ; __vbaSetSystemError
  loc_00662727: mov edx, var_28
  loc_0066272A: mov esi, [00401190h] ; __vbaStrToUnicode
  loc_00662730: push edx
  loc_00662731: push edi
  loc_00662732: call __vbaStrToUnicode
  loc_00662734: mov eax, var_34
  loc_00662737: mov ecx, strkeyvalue
  loc_0066273A: push eax
  loc_0066273B: push ecx
  loc_0066273C: call __vbaStrToUnicode
  loc_0066273E: mov edx, var_38
  loc_00662741: mov eax, strfullpath
  loc_00662744: push edx
  loc_00662745: push eax
  loc_00662746: call __vbaStrToUnicode
  loc_00662748: lea ecx, var_3C
  loc_0066274B: lea edx, var_38
  loc_0066274E: push ecx
  loc_0066274F: push edx
  loc_00662750: lea eax, var_34
  loc_00662753: lea ecx, var_30
  loc_00662756: push eax
  loc_00662757: lea edx, var_2C
  loc_0066275A: push ecx
  loc_0066275B: lea eax, var_28
  loc_0066275E: push edx
  loc_0066275F: push eax
  loc_00662760: push 00000006h
  loc_00662762: call [00401204h] ; __vbaFreeStrList
  loc_00662768: add esp, 0000001Ch
  loc_0066276B: push 006627A6h
  loc_00662770: jmp 006627A5h
  loc_00662772: test var_4, 04h
  loc_00662776: jz 00662781h
  loc_00662778: lea ecx, var_24
  loc_0066277B: call [00401020h] ; __vbaFreeVar
  loc_00662781: lea ecx, var_3C
  loc_00662784: lea edx, var_38
  loc_00662787: push ecx
  loc_00662788: lea eax, var_34
  loc_0066278B: push edx
  loc_0066278C: lea ecx, var_30
  loc_0066278F: push eax
  loc_00662790: lea edx, var_2C
  loc_00662793: push ecx
  loc_00662794: lea eax, var_28
  loc_00662797: push edx
  loc_00662798: push eax
  loc_00662799: push 00000006h
  loc_0066279B: call [00401204h] ; __vbaFreeStrList
  loc_006627A1: add esp, 0000001Ch
  loc_006627A4: ret
  loc_006627A5: ret
  loc_006627A6: mov eax, Me
  loc_006627A9: push eax
  loc_006627AA: mov ecx, [eax]
  loc_006627AC: call [ecx+00000008h]
  loc_006627AF: mov edx, arg_1C
  loc_006627B2: mov eax, var_24
  loc_006627B5: mov ecx, var_20
  loc_006627B8: mov [edx], eax
  loc_006627BA: mov eax, var_1C
  loc_006627BD: mov [edx+00000004h], ecx
  loc_006627C0: mov ecx, var_18
  loc_006627C3: mov [edx+00000008h], eax
  loc_006627C6: mov [edx+0000000Ch], ecx
  loc_006627C9: mov eax, var_4
  loc_006627CC: mov ecx, var_14
  loc_006627CF: pop edi
  loc_006627D0: pop esi
  loc_006627D1: mov fs:[00000000h], ecx
  loc_006627D8: pop ebx
  loc_006627D9: mov esp, ebp
  loc_006627DB: pop ebp
  loc_006627DC: retn 0018h
End Function

Public Sub HackerScan() '6627E0
  loc_006627E0: push ebp
  loc_006627E1: mov ebp, esp
  loc_006627E3: sub esp, 0000000Ch
  loc_006627E6: push 00405696h ; __vbaExceptHandler
  loc_006627EB: mov eax, fs:[00000000h]
  loc_006627F1: push eax
  loc_006627F2: mov fs:[00000000h], esp
  loc_006627F9: sub esp, 000000B8h
  loc_006627FF: push ebx
  loc_00662800: push esi
  loc_00662801: push edi
  loc_00662802: mov var_C, esp
  loc_00662805: mov var_8, 00405558h
  loc_0066280C: xor esi, esi
  loc_0066280E: mov var_4, esi
  loc_00662811: mov eax, Me
  loc_00662814: push eax
  loc_00662815: mov ecx, [eax]
  loc_00662817: call [ecx+00000004h]
  loc_0066281A: mov edi, [00401070h] ; __vbaStrCat
  loc_00662820: push 00434A94h
  loc_00662825: push 0043FEF4h
  loc_0066282A: mov var_1C, esi
  loc_0066282D: mov var_24, esi
  loc_00662830: mov var_28, esi
  loc_00662833: mov var_2C, esi
  loc_00662836: mov var_30, esi
  loc_00662839: mov var_34, esi
  loc_0066283C: mov var_38, esi
  loc_0066283F: mov var_48, esi
  loc_00662842: mov var_58, esi
  loc_00662845: mov var_68, esi
  loc_00662848: mov var_78, esi
  loc_0066284B: mov var_88, esi
  loc_00662851: mov var_BC, esi
  loc_00662857: call edi
  loc_00662859: mov ebx, [0040126Ch] ; __vbaStrMove
  loc_0066285F: mov edx, eax
  loc_00662861: lea ecx, var_24
  loc_00662864: call ebx
  loc_00662866: push eax
  loc_00662867: push 0043FEFCh
  loc_0066286C: call edi
  loc_0066286E: mov edx, eax
  loc_00662870: lea ecx, var_28
  loc_00662873: call ebx
  loc_00662875: push eax
  loc_00662876: push 0043FF04h
  loc_0066287B: call edi
  loc_0066287D: mov edx, eax
  loc_0066287F: lea ecx, var_2C
  loc_00662882: call ebx
  loc_00662884: push eax
  loc_00662885: push 0043FF0Ch
  loc_0066288A: call edi
  loc_0066288C: mov edx, eax
  loc_0066288E: lea ecx, var_30
  loc_00662891: call ebx
  loc_00662893: push eax
  loc_00662894: push 0042F350h ; "-"
  loc_00662899: call edi
  loc_0066289B: mov edx, eax
  loc_0066289D: lea ecx, var_34
  loc_006628A0: call ebx
  loc_006628A2: push eax
  loc_006628A3: push 0043FEFCh
  loc_006628A8: call edi
  loc_006628AA: mov edx, eax
  loc_006628AC: lea ecx, var_38
  loc_006628AF: call ebx
  loc_006628B1: push eax
  loc_006628B2: push 00434A94h
  loc_006628B7: call edi
  loc_006628B9: mov edx, eax
  loc_006628BB: lea ecx, var_1C
  loc_006628BE: call ebx
  loc_006628C0: lea edx, var_38
  loc_006628C3: lea eax, var_34
  loc_006628C6: push edx
  loc_006628C7: lea ecx, var_30
  loc_006628CA: push eax
  loc_006628CB: lea edx, var_2C
  loc_006628CE: push ecx
  loc_006628CF: lea eax, var_28
  loc_006628D2: push edx
  loc_006628D3: lea ecx, var_24
  loc_006628D6: push eax
  loc_006628D7: push ecx
  loc_006628D8: push 00000006h
  loc_006628DA: call [00401204h] ; __vbaFreeStrList
  loc_006628E0: add esp, 0000001Ch
  loc_006628E3: lea edx, var_BC
  loc_006628E9: lea eax, var_24
  loc_006628EC: mov var_BC, esi
  loc_006628F2: push esi
  loc_006628F3: push 00000080h
  loc_006628F8: push 00000003h
  loc_006628FA: push edx
  loc_006628FB: push 00000003h
  loc_006628FD: push C0000000h
  loc_00662902: push 0043FF14h ; "\\.\SICE"
  loc_00662907: push eax
  loc_00662908: call [00401230h] ; __vbaStrToAnsi
  loc_0066290E: push eax
  loc_0066290F: CreateFile(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_00662914: mov edi, eax
  loc_00662916: call [0040107Ch] ; __vbaSetSystemError
  loc_0066291C: lea ecx, var_24
  loc_0066291F: call [0040129Ch] ; __vbaFreeStr
  loc_00662925: cmp edi, FFFFFFFFh
  loc_00662928: jz 0066296Ch
  loc_0066292A: push esi
  loc_0066292B: push 00000080h
  loc_00662930: lea ecx, var_BC
  loc_00662936: push 00000003h
  loc_00662938: push ecx
  loc_00662939: push 00000003h
  loc_0066293B: push C0000000h
  loc_00662940: lea edx, var_24
  loc_00662943: push 0043FF2Ch ; "\\.\NTICE"
  loc_00662948: push edx
  loc_00662949: mov var_BC, esi
  loc_0066294F: call [00401230h] ; __vbaStrToAnsi
  loc_00662955: push eax
  loc_00662956: CreateFile(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_0066295B: mov edi, eax
  loc_0066295D: call [0040107Ch] ; __vbaSetSystemError
  loc_00662963: lea ecx, var_24
  loc_00662966: call [0040129Ch] ; __vbaFreeStr
  loc_0066296C: mov eax, var_1C
  loc_0066296F: push esi
  loc_00662970: lea ecx, var_24
  loc_00662973: push eax
  loc_00662974: push ecx
  loc_00662975: call [00401230h] ; __vbaStrToAnsi
  loc_0066297B: push eax
  loc_0066297C: FindWindow(%x1v, %x2v)
  loc_00662981: mov ebx, eax
  loc_00662983: call [0040107Ch] ; __vbaSetSystemError
  loc_00662989: mov edx, var_24
  loc_0066298C: lea eax, var_1C
  loc_0066298F: push edx
  loc_00662990: push eax
  loc_00662991: call [00401190h] ; __vbaStrToUnicode
  loc_00662997: xor ecx, ecx
  loc_00662999: cmp ebx, esi
  loc_0066299B: setnz cl
  loc_0066299E: neg ecx
  loc_006629A0: xor edx, edx
  loc_006629A2: cmp edi, FFFFFFFFh
  loc_006629A5: setnz dl
  loc_006629A8: neg edx
  loc_006629AA: or ecx, edx
  loc_006629AC: mov ebx, ecx
  loc_006629AE: lea ecx, var_24
  loc_006629B1: call [0040129Ch] ; __vbaFreeStr
  loc_006629B7: cmp bx, si
  loc_006629BA: jz 00662A30h
  loc_006629BC: mov ecx, 80020004h
  loc_006629C1: mov eax, 0000000Ah
  loc_006629C6: mov var_70, ecx
  loc_006629C9: mov var_60, ecx
  loc_006629CC: mov var_50, ecx
  loc_006629CF: lea edx, var_88
  loc_006629D5: lea ecx, var_48
  loc_006629D8: mov var_78, eax
  loc_006629DB: mov var_68, eax
  loc_006629DE: mov var_58, eax
  loc_006629E1: mov var_80, 0043FF44h ; "Hacking activity detected! Application will exit"
  loc_006629E8: mov var_88, 00000008h
  loc_006629F2: call [00401238h] ; __vbaVarDup
  loc_006629F8: lea eax, var_78
  loc_006629FB: lea ecx, var_68
  loc_006629FE: push eax
  loc_006629FF: lea edx, var_58
  loc_00662A02: push ecx
  loc_00662A03: push edx
  loc_00662A04: lea eax, var_48
  loc_00662A07: push esi
  loc_00662A08: push eax
  loc_00662A09: call [004010B4h] ; rtcMsgBox
  loc_00662A0F: lea ecx, var_78
  loc_00662A12: lea edx, var_68
  loc_00662A15: push ecx
  loc_00662A16: lea eax, var_58
  loc_00662A19: push edx
  loc_00662A1A: lea ecx, var_48
  loc_00662A1D: push eax
  loc_00662A1E: push ecx
  loc_00662A1F: push 00000004h
  loc_00662A21: call [0040103Ch] ; __vbaFreeVarList
  loc_00662A27: add esp, 00000014h
  loc_00662A2A: call [00401038h] ; __vbaEnd
  loc_00662A30: cmp edi, FFFFFFFFh
  loc_00662A33: jz 00662A41h
  loc_00662A35: push edi
  loc_00662A36: CloseHandle(%x1v)
  loc_00662A3B: call [0040107Ch] ; __vbaSetSystemError
  loc_00662A41: push 00662A8Eh
  loc_00662A46: jmp 00662A84h
  loc_00662A48: lea edx, var_38
  loc_00662A4B: lea eax, var_34
  loc_00662A4E: push edx
  loc_00662A4F: lea ecx, var_30
  loc_00662A52: push eax
  loc_00662A53: lea edx, var_2C
  loc_00662A56: push ecx
  loc_00662A57: lea eax, var_28
  loc_00662A5A: push edx
  loc_00662A5B: lea ecx, var_24
  loc_00662A5E: push eax
  loc_00662A5F: push ecx
  loc_00662A60: push 00000006h
  loc_00662A62: call [00401204h] ; __vbaFreeStrList
  loc_00662A68: lea edx, var_78
  loc_00662A6B: lea eax, var_68
  loc_00662A6E: push edx
  loc_00662A6F: lea ecx, var_58
  loc_00662A72: push eax
  loc_00662A73: lea edx, var_48
  loc_00662A76: push ecx
  loc_00662A77: push edx
  loc_00662A78: push 00000004h
  loc_00662A7A: call [0040103Ch] ; __vbaFreeVarList
  loc_00662A80: add esp, 00000030h
  loc_00662A83: ret
  loc_00662A84: lea ecx, var_1C
  loc_00662A87: call [0040129Ch] ; __vbaFreeStr
  loc_00662A8D: ret
  loc_00662A8E: mov eax, Me
  loc_00662A91: push eax
  loc_00662A92: mov ecx, [eax]
  loc_00662A94: call [ecx+00000008h]
  loc_00662A97: mov eax, var_4
  loc_00662A9A: mov ecx, var_14
  loc_00662A9D: pop edi
  loc_00662A9E: pop esi
  loc_00662A9F: mov fs:[00000000h], ecx
  loc_00662AA6: pop ebx
  loc_00662AA7: mov esp, ebp
  loc_00662AA9: pop ebp
  loc_00662AAA: retn 0004h
End Sub

Private Sub Proc_74_11_65E390
  loc_0065E390: push ebp
  loc_0065E391: mov ebp, esp
  loc_0065E393: sub esp, 00000008h
  loc_0065E396: push 00405696h ; __vbaExceptHandler
  loc_0065E39B: mov eax, fs:[00000000h]
  loc_0065E3A1: push eax
  loc_0065E3A2: mov fs:[00000000h], esp
  loc_0065E3A9: sub esp, 00000200h
  loc_0065E3AF: push ebx
  loc_0065E3B0: push esi
  loc_0065E3B1: push edi
  loc_0065E3B2: mov var_8, esp
  loc_0065E3B5: mov var_4, 00405470h
  loc_0065E3BC: mov esi, [00401130h] ; __vbaAryConstruct2
  loc_0065E3C2: push 00000008h
  loc_0065E3C4: lea eax, var_38
  loc_0065E3C7: xor ebx, ebx
  loc_0065E3C9: push 0043FA14h
  loc_0065E3CE: push eax
  loc_0065E3CF: mov var_20, ebx
  loc_0065E3D2: mov var_A4, ebx
  loc_0065E3D8: mov var_A8, ebx
  loc_0065E3DE: mov var_AC, ebx
  loc_0065E3E4: mov var_B0, ebx
  loc_0065E3EA: mov var_B4, ebx
  loc_0065E3F0: mov var_B8, ebx
  loc_0065E3F6: mov var_C8, ebx
  loc_0065E3FC: mov var_D8, ebx
  loc_0065E402: mov var_E8, ebx
  loc_0065E408: mov var_F8, ebx
  loc_0065E40E: mov var_108, ebx
  loc_0065E414: mov var_118, ebx
  loc_0065E41A: mov var_128, ebx
  loc_0065E420: mov var_138, ebx
  loc_0065E426: mov var_148, ebx
  loc_0065E42C: mov var_158, ebx
  loc_0065E432: mov var_168, ebx
  loc_0065E438: mov var_188, ebx
  loc_0065E43E: mov var_18C, ebx
  loc_0065E444: mov var_190, ebx
  loc_0065E44A: mov var_194, ebx
  loc_0065E450: mov var_198, ebx
  loc_0065E456: mov var_1B8, ebx
  loc_0065E45C: mov var_1C8, ebx
  loc_0065E462: mov var_1D8, ebx
  loc_0065E468: mov var_1E8, ebx
  loc_0065E46E: mov var_1F8, ebx
  loc_0065E474: call __vbaAryConstruct2
  loc_0065E476: push 00000008h
  loc_0065E478: lea ecx, var_58
  loc_0065E47B: push 0043FA14h
  loc_0065E480: push ecx
  loc_0065E481: call __vbaAryConstruct2
  loc_0065E483: push 00000008h
  loc_0065E485: lea edx, var_74
  loc_0065E488: push 0043FA14h
  loc_0065E48D: push edx
  loc_0065E48E: call __vbaAryConstruct2
  loc_0065E490: push 0043F9B4h
  loc_0065E495: lea eax, var_90
  loc_0065E49B: push 0043FA30h
  loc_0065E4A0: push eax
  loc_0065E4A1: call __vbaAryConstruct2
  loc_0065E4A3: mov ecx, var_2C
  loc_0065E4A6: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0065E4AC: mov edx, 0043F8C0h ; "Win32_Processor,ProcessorId"
  loc_0065E4B1: add ecx, 00000004h
  loc_0065E4B4: call __vbaStrCopy
  loc_0065E4B6: mov eax, var_2C
  loc_0065E4B9: mov edx, 0043F8FCh ; "Win32_OperatingSystem,SerialNumber"
  loc_0065E4BE: lea ecx, [eax+00000008h]
  loc_0065E4C1: call __vbaStrCopy
  loc_0065E4C3: mov eax, 00000001h
  loc_0065E4C8: lea ecx, var_138
  loc_0065E4CE: mov var_130, eax
  loc_0065E4D4: mov var_150, eax
  loc_0065E4DA: lea edx, var_148
  loc_0065E4E0: push ecx
  loc_0065E4E1: lea eax, var_158
  loc_0065E4E7: mov edi, 00000002h
  loc_0065E4EC: push edx
  loc_0065E4ED: lea ecx, var_1D8
  loc_0065E4F3: push eax
  loc_0065E4F4: mov var_138, edi
  loc_0065E4FA: mov var_140, edi
  loc_0065E500: mov var_148, edi
  loc_0065E506: mov var_158, edi
  loc_0065E50C: push ecx
  loc_0065E50D: lea edx, var_1C8
  loc_0065E513: lea eax, var_20
  loc_0065E516: push edx
  loc_0065E517: push eax
  loc_0065E518: call [004010A0h] ; __vbaVarForInit
  loc_0065E51E: mov esi, [0040121Ch] ; __vbaI4Var
  loc_0065E524: cmp eax, ebx
  loc_0065E526: jz 0065EBCCh
  loc_0065E52C: lea edx, var_138
  loc_0065E532: lea ecx, var_C8
  loc_0065E538: mov var_140, ebx
  loc_0065E53E: mov var_148, edi
  loc_0065E544: mov var_130, 00433074h
  loc_0065E54E: mov var_138, 00000008h
  loc_0065E558: call [00401238h] ; __vbaVarDup
  loc_0065E55E: lea ecx, var_20
  loc_0065E561: push ecx
  loc_0065E562: call __vbaI4Var
  loc_0065E564: mov edi, eax
  loc_0065E566: cmp edi, 00000003h
  loc_0065E569: jb 0065E571h
  loc_0065E56B: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E571: mov eax, var_2C
  loc_0065E574: push ebx
  loc_0065E575: lea edx, var_C8
  loc_0065E57B: push FFFFFFFFh
  loc_0065E57D: mov ecx, [eax+edi*4]
  loc_0065E580: push edx
  loc_0065E581: lea edx, var_D8
  loc_0065E587: push ecx
  loc_0065E588: push edx
  loc_0065E589: call [0040118Ch] ; rtcSplit
  loc_0065E58F: mov ecx, var_148
  loc_0065E595: mov edx, var_144
  loc_0065E59B: sub esp, 00000010h
  loc_0065E59E: mov eax, esp
  loc_0065E5A0: push 00000001h
  loc_0065E5A2: mov [eax], ecx
  loc_0065E5A4: mov ecx, var_140
  loc_0065E5AA: mov [eax+00000004h], edx
  loc_0065E5AD: mov edx, var_13C
  loc_0065E5B3: mov [eax+00000008h], ecx
  loc_0065E5B6: lea ecx, var_118
  loc_0065E5BC: mov [eax+0000000Ch], edx
  loc_0065E5BF: lea eax, var_D8
  loc_0065E5C5: push eax
  loc_0065E5C6: push ecx
  loc_0065E5C7: call [00401054h] ; __vbaVarIndexLoadRef
  loc_0065E5CD: add esp, 0000001Ch
  loc_0065E5D0: lea edx, var_168
  loc_0065E5D6: lea ecx, var_E8
  loc_0065E5DC: mov ebx, eax
  loc_0065E5DE: mov var_F0, 80020004h
  loc_0065E5E8: mov var_F8, 0000000Ah
  loc_0065E5F2: mov var_160, 0043F948h ; "winmgmts:{impersonationLevel=impersonate}"
  loc_0065E5FC: mov var_168, 00000008h
  loc_0065E606: call [00401238h] ; __vbaVarDup
  loc_0065E60C: lea edx, var_F8
  loc_0065E612: lea eax, var_E8
  loc_0065E618: push edx
  loc_0065E619: lea ecx, var_108
  loc_0065E61F: push eax
  loc_0065E620: push ecx
  loc_0065E621: call [00401064h] ; rtcGetObject
  loc_0065E627: lea edx, var_20
  loc_0065E62A: push edx
  loc_0065E62B: call __vbaI4Var
  loc_0065E62D: mov edi, eax
  loc_0065E62F: cmp edi, 00000003h
  loc_0065E632: jb 0065E63Ah
  loc_0065E634: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E63A: push 0043F9B4h
  loc_0065E63F: mov edx, var_184
  loc_0065E645: sub esp, 00000010h
  loc_0065E648: mov eax, 0000400Ch
  loc_0065E64D: mov ecx, esp
  loc_0065E64F: push 00000001h
  loc_0065E651: push 0043F99Ch ; "InstancesOf"
  loc_0065E656: mov [ecx], eax
  loc_0065E658: mov eax, var_17C
  loc_0065E65E: mov [ecx+00000004h], edx
  loc_0065E661: lea edx, var_128
  loc_0065E667: mov [ecx+00000008h], ebx
  loc_0065E66A: mov [ecx+0000000Ch], eax
  loc_0065E66D: lea ecx, var_108
  loc_0065E673: push ecx
  loc_0065E674: push edx
  loc_0065E675: call [00401248h] ; __vbaVarLateMemCallLd
  loc_0065E67B: add esp, 00000020h
  loc_0065E67E: push eax
  loc_0065E67F: call [00401144h] ; __vbaCastObjVar
  loc_0065E685: push eax
  loc_0065E686: lea eax, var_B4
  loc_0065E68C: push eax
  loc_0065E68D: call [004010B8h] ; __vbaObjSet
  loc_0065E693: mov ecx, var_84
  loc_0065E699: push eax
  loc_0065E69A: lea edx, [ecx+edi*4]
  loc_0065E69D: push edx
  loc_0065E69E: call [004010C8h] ; __vbaObjSetAddref
  loc_0065E6A4: lea ecx, var_B4
  loc_0065E6AA: call [004012A0h] ; __vbaFreeObj
  loc_0065E6B0: lea eax, var_128
  loc_0065E6B6: lea ecx, var_108
  loc_0065E6BC: push eax
  loc_0065E6BD: lea edx, var_118
  loc_0065E6C3: push ecx
  loc_0065E6C4: lea eax, var_F8
  loc_0065E6CA: push edx
  loc_0065E6CB: lea ecx, var_E8
  loc_0065E6D1: push eax
  loc_0065E6D2: lea edx, var_D8
  loc_0065E6D8: push ecx
  loc_0065E6D9: lea eax, var_C8
  loc_0065E6DF: push edx
  loc_0065E6E0: push eax
  loc_0065E6E1: push 00000007h
  loc_0065E6E3: call [0040103Ch] ; __vbaFreeVarList
  loc_0065E6E9: add esp, 00000020h
  loc_0065E6EC: lea ecx, var_20
  loc_0065E6EF: push ecx
  loc_0065E6F0: call __vbaI4Var
  loc_0065E6F2: mov edi, eax
  loc_0065E6F4: cmp edi, 00000003h
  loc_0065E6F7: jb 0065E6FFh
  loc_0065E6F9: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E6FF: mov eax, var_68
  loc_0065E702: mov ebx, [004011FCh] ; __vbaStrCopy
  loc_0065E708: mov edx, 0042DFECh
  loc_0065E70D: lea ecx, [eax+edi*4]
  loc_0065E710: call ebx
  loc_0065E712: lea ecx, var_20
  loc_0065E715: push ecx
  loc_0065E716: call __vbaI4Var
  loc_0065E718: mov edi, eax
  loc_0065E71A: cmp edi, 00000003h
  loc_0065E71D: jb 0065E725h
  loc_0065E71F: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E725: mov edx, var_84
  loc_0065E72B: lea ecx, var_A8
  loc_0065E731: mov eax, [edx+edi*4]
  loc_0065E734: lea edx, var_1B8
  loc_0065E73A: push eax
  loc_0065E73B: push ecx
  loc_0065E73C: push edx
  loc_0065E73D: push 0043F9C4h
  loc_0065E742: call [00401098h] ; __vbaForEachCollObj
  loc_0065E748: test eax, eax
  loc_0065E74A: jz 0065EADBh
  loc_0065E750: mov edx, 0042DFECh
  loc_0065E755: lea ecx, var_A4
  loc_0065E75B: call ebx
  loc_0065E75D: lea eax, var_20
  loc_0065E760: push eax
  loc_0065E761: call __vbaI4Var
  loc_0065E763: mov edi, eax
  loc_0065E765: cmp edi, 00000003h
  loc_0065E768: jb 0065E770h
  loc_0065E76A: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E770: mov ecx, var_4C
  loc_0065E773: mov edx, 0042DFECh
  loc_0065E778: lea ecx, [ecx+edi*4]
  loc_0065E77B: call ebx
  loc_0065E77D: lea edx, var_138
  loc_0065E783: lea ecx, var_C8
  loc_0065E789: mov var_140, 00000001h
  loc_0065E793: mov var_148, 00000002h
  loc_0065E79D: mov var_130, 00433074h
  loc_0065E7A7: mov var_138, 00000008h
  loc_0065E7B1: call [00401238h] ; __vbaVarDup
  loc_0065E7B7: lea edx, var_20
  loc_0065E7BA: push edx
  loc_0065E7BB: call __vbaI4Var
  loc_0065E7BD: mov edi, eax
  loc_0065E7BF: cmp edi, 00000003h
  loc_0065E7C2: jb 0065E7CAh
  loc_0065E7C4: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E7CA: mov ecx, var_2C
  loc_0065E7CD: push 00000000h
  loc_0065E7CF: lea eax, var_C8
  loc_0065E7D5: push FFFFFFFFh
  loc_0065E7D7: mov edx, [ecx+edi*4]
  loc_0065E7DA: push eax
  loc_0065E7DB: lea eax, var_D8
  loc_0065E7E1: push edx
  loc_0065E7E2: push eax
  loc_0065E7E3: call [0040118Ch] ; rtcSplit
  loc_0065E7E9: mov edx, var_148
  loc_0065E7EF: mov eax, var_144
  loc_0065E7F5: sub esp, 00000010h
  loc_0065E7F8: mov ecx, esp
  loc_0065E7FA: push 00000001h
  loc_0065E7FC: mov [ecx], edx
  loc_0065E7FE: mov edx, var_140
  loc_0065E804: mov [ecx+00000004h], eax
  loc_0065E807: mov eax, var_13C
  loc_0065E80D: mov [ecx+00000008h], edx
  loc_0065E810: lea edx, var_E8
  loc_0065E816: mov [ecx+0000000Ch], eax
  loc_0065E819: lea ecx, var_D8
  loc_0065E81F: push ecx
  loc_0065E820: push edx
  loc_0065E821: call [004010D0h] ; __vbaVarIndexLoad
  loc_0065E827: mov eax, var_A8
  loc_0065E82D: add esp, 0000001Ch
  loc_0065E830: lea edx, var_B4
  loc_0065E836: mov ecx, [eax]
  loc_0065E838: push edx
  loc_0065E839: push eax
  loc_0065E83A: call [ecx+0000006Ch]
  loc_0065E83D: test eax, eax
  loc_0065E83F: fnclex
  loc_0065E841: jge 0065E858h
  loc_0065E843: mov ecx, var_A8
  loc_0065E849: push 0000006Ch
  loc_0065E84B: push 0043F9C4h
  loc_0065E850: push ecx
  loc_0065E851: push eax
  loc_0065E852: call [00401080h] ; __vbaHresultCheckObj
  loc_0065E858: mov edi, var_B4
  loc_0065E85E: lea edx, var_B8
  loc_0065E864: push edx
  loc_0065E865: lea eax, var_E8
  loc_0065E86B: mov ebx, [edi]
  loc_0065E86D: push 00000000h
  loc_0065E86F: lea ecx, var_AC
  loc_0065E875: push eax
  loc_0065E876: push ecx
  loc_0065E877: call [004011C0h] ; __vbaStrVarVal
  loc_0065E87D: push eax
  loc_0065E87E: push edi
  loc_0065E87F: call [ebx+00000020h]
  loc_0065E882: test eax, eax
  loc_0065E884: fnclex
  loc_0065E886: jge 0065E897h
  loc_0065E888: push 00000020h
  loc_0065E88A: push 0043F9D4h
  loc_0065E88F: push edi
  loc_0065E890: push eax
  loc_0065E891: call [00401080h] ; __vbaHresultCheckObj
  loc_0065E897: mov eax, var_B8
  loc_0065E89D: lea ecx, var_F8
  loc_0065E8A3: push ecx
  loc_0065E8A4: push eax
  loc_0065E8A5: mov edx, [eax]
  loc_0065E8A7: mov edi, eax
  loc_0065E8A9: call [edx+0000001Ch]
  loc_0065E8AC: test eax, eax
  loc_0065E8AE: fnclex
  loc_0065E8B0: jge 0065E8C1h
  loc_0065E8B2: push 0000001Ch
  loc_0065E8B4: push 0043F9E4h
  loc_0065E8B9: push edi
  loc_0065E8BA: push eax
  loc_0065E8BB: call [00401080h] ; __vbaHresultCheckObj
  loc_0065E8C1: lea edx, var_20
  loc_0065E8C4: push edx
  loc_0065E8C5: call __vbaI4Var
  loc_0065E8C7: mov edi, eax
  loc_0065E8C9: cmp edi, 00000003h
  loc_0065E8CC: jb 0065E8D4h
  loc_0065E8CE: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E8D4: lea eax, var_F8
  loc_0065E8DA: push eax
  loc_0065E8DB: call [00401034h] ; __vbaStrVarMove
  loc_0065E8E1: mov edx, eax
  loc_0065E8E3: lea ecx, var_B0
  loc_0065E8E9: call [0040126Ch] ; __vbaStrMove
  loc_0065E8EF: mov ecx, var_68
  loc_0065E8F2: mov edx, eax
  loc_0065E8F4: lea ecx, [ecx+edi*4]
  loc_0065E8F7: call [004011FCh] ; __vbaStrCopy
  loc_0065E8FD: lea edx, var_B0
  loc_0065E903: lea eax, var_AC
  loc_0065E909: push edx
  loc_0065E90A: push eax
  loc_0065E90B: push 00000002h
  loc_0065E90D: call [00401204h] ; __vbaFreeStrList
  loc_0065E913: lea ecx, var_B8
  loc_0065E919: lea edx, var_B4
  loc_0065E91F: push ecx
  loc_0065E920: push edx
  loc_0065E921: push 00000002h
  loc_0065E923: call [0040104Ch] ; __vbaFreeObjList
  loc_0065E929: lea eax, var_F8
  loc_0065E92F: lea ecx, var_E8
  loc_0065E935: push eax
  loc_0065E936: lea edx, var_D8
  loc_0065E93C: push ecx
  loc_0065E93D: lea eax, var_C8
  loc_0065E943: push edx
  loc_0065E944: push eax
  loc_0065E945: push 00000004h
  loc_0065E947: call [0040103Ch] ; __vbaFreeVarList
  loc_0065E94D: add esp, 0000002Ch
  loc_0065E950: lea ecx, var_20
  loc_0065E953: push ecx
  loc_0065E954: call __vbaI4Var
  loc_0065E956: mov edi, eax
  loc_0065E958: cmp edi, 00000003h
  loc_0065E95B: jb 0065E963h
  loc_0065E95D: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E963: mov edx, var_68
  loc_0065E966: mov eax, [edx+edi*4]
  loc_0065E969: push eax
  loc_0065E96A: call [00401030h] ; __vbaLenBstr
  loc_0065E970: mov ecx, eax
  loc_0065E972: call [00401138h] ; __vbaI2I4
  loc_0065E978: mov ebx, 00000001h
  loc_0065E97D: mov var_200, eax
  loc_0065E983: mov var_40, ebx
  loc_0065E986: cmp bx, var_200
  loc_0065E98D: jg 0065EAB7h
  loc_0065E993: lea ecx, var_20
  loc_0065E996: mov var_C0, 00000001h
  loc_0065E9A0: push ecx
  loc_0065E9A1: mov var_C8, 00000002h
  loc_0065E9AB: call __vbaI4Var
  loc_0065E9AD: mov edi, eax
  loc_0065E9AF: cmp edi, 00000003h
  loc_0065E9B2: jb 0065E9BAh
  loc_0065E9B4: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065E9BA: mov edx, var_68
  loc_0065E9BD: lea ecx, var_C8
  loc_0065E9C3: push ecx
  loc_0065E9C4: lea ecx, var_D8
  loc_0065E9CA: lea eax, [edx+edi*4]
  loc_0065E9CD: mov var_138, 00004008h
  loc_0065E9D7: movsx edx, bx
  loc_0065E9DA: mov var_130, eax
  loc_0065E9E0: lea eax, var_138
  loc_0065E9E6: push edx
  loc_0065E9E7: push eax
  loc_0065E9E8: push ecx
  loc_0065E9E9: call [00401100h] ; rtcMidCharVar
  loc_0065E9EF: lea edx, var_D8
  loc_0065E9F5: push edx
  loc_0065E9F6: call [00401034h] ; __vbaStrVarMove
  loc_0065E9FC: mov edx, eax
  loc_0065E9FE: lea ecx, var_A4
  loc_0065EA04: call [0040126Ch] ; __vbaStrMove
  loc_0065EA0A: lea eax, var_D8
  loc_0065EA10: lea ecx, var_C8
  loc_0065EA16: push eax
  loc_0065EA17: push ecx
  loc_0065EA18: push 00000002h
  loc_0065EA1A: call [0040103Ch] ; __vbaFreeVarList
  loc_0065EA20: mov edx, var_A4
  loc_0065EA26: add esp, 0000000Ch
  loc_0065EA29: push edx
  loc_0065EA2A: push 0043F9F8h ; "[0-9A-Fa-f]"
  loc_0065EA2F: call [004010ACh] ; __vbaStrLike
  loc_0065EA35: test ax, ax
  loc_0065EA38: jz 0065EA9Fh
  loc_0065EA3A: lea eax, var_20
  loc_0065EA3D: push eax
  loc_0065EA3E: call __vbaI4Var
  loc_0065EA40: mov ebx, eax
  loc_0065EA42: cmp ebx, 00000003h
  loc_0065EA45: jb 0065EA4Dh
  loc_0065EA47: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EA4D: lea ecx, var_20
  loc_0065EA50: push ecx
  loc_0065EA51: call __vbaI4Var
  loc_0065EA53: mov edi, eax
  loc_0065EA55: cmp edi, 00000003h
  loc_0065EA58: jb 0065EA60h
  loc_0065EA5A: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EA60: mov edx, var_4C
  loc_0065EA63: mov ecx, var_A4
  loc_0065EA69: mov eax, [edx+ebx*4]
  loc_0065EA6C: push eax
  loc_0065EA6D: push ecx
  loc_0065EA6E: call [00401070h] ; __vbaStrCat
  loc_0065EA74: mov edx, eax
  loc_0065EA76: lea ecx, var_AC
  loc_0065EA7C: call [0040126Ch] ; __vbaStrMove
  loc_0065EA82: mov edx, eax
  loc_0065EA84: mov eax, var_4C
  loc_0065EA87: lea ecx, [eax+edi*4]
  loc_0065EA8A: call [004011FCh] ; __vbaStrCopy
  loc_0065EA90: lea ecx, var_AC
  loc_0065EA96: call [0040129Ch] ; __vbaFreeStr
  loc_0065EA9C: mov ebx, var_40
  loc_0065EA9F: mov eax, 00000001h
  loc_0065EAA4: add ax, bx
  loc_0065EAA7: jo 0065EECEh
  loc_0065EAAD: mov var_40, eax
  loc_0065EAB0: mov ebx, eax
  loc_0065EAB2: jmp 0065E986h
  loc_0065EAB7: lea ecx, var_A8
  loc_0065EABD: lea edx, var_1B8
  loc_0065EAC3: push ecx
  loc_0065EAC4: push edx
  loc_0065EAC5: push 0043F9C4h
  loc_0065EACA: call [00401104h] ; __vbaNextEachCollObj
  loc_0065EAD0: mov ebx, [004011FCh] ; __vbaStrCopy
  loc_0065EAD6: jmp 0065E748h
  loc_0065EADB: lea eax, var_20
  loc_0065EADE: push eax
  loc_0065EADF: call __vbaI4Var
  loc_0065EAE1: cmp eax, 00000003h
  loc_0065EAE4: mov var_1A0, eax
  loc_0065EAEA: jb 0065EAF2h
  loc_0065EAEC: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EAF2: lea ecx, var_20
  loc_0065EAF5: push ecx
  loc_0065EAF6: call __vbaI4Var
  loc_0065EAF8: mov edi, eax
  loc_0065EAFA: cmp edi, 00000003h
  loc_0065EAFD: jb 0065EB05h
  loc_0065EAFF: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EB05: mov edx, var_4C
  loc_0065EB08: mov eax, var_1A0
  loc_0065EB0E: mov ecx, var_68
  loc_0065EB11: mov edx, [edx+eax*4]
  loc_0065EB14: lea ecx, [ecx+edi*4]
  loc_0065EB17: call ebx
  loc_0065EB19: lea edx, var_20
  loc_0065EB1C: push edx
  loc_0065EB1D: call __vbaI4Var
  loc_0065EB1F: mov edi, eax
  loc_0065EB21: cmp edi, 00000003h
  loc_0065EB24: jb 0065EB2Ch
  loc_0065EB26: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EB2C: mov eax, var_68
  loc_0065EB2F: lea edx, var_138
  loc_0065EB35: push 00000004h
  loc_0065EB37: push edx
  loc_0065EB38: lea ecx, [eax+edi*4]
  loc_0065EB3B: lea eax, var_C8
  loc_0065EB41: push eax
  loc_0065EB42: mov var_130, ecx
  loc_0065EB48: mov var_138, 00004008h
  loc_0065EB52: call [00401274h] ; rtcRightCharVar
  loc_0065EB58: lea ecx, var_20
  loc_0065EB5B: push ecx
  loc_0065EB5C: call __vbaI4Var
  loc_0065EB5E: mov edi, eax
  loc_0065EB60: cmp edi, 00000003h
  loc_0065EB63: jb 0065EB6Bh
  loc_0065EB65: call [00401120h] ; __vbaGenerateBoundsError
  loc_0065EB6B: lea edx, var_C8
  loc_0065EB71: push edx
  loc_0065EB72: call [00401034h] ; __vbaStrVarMove
  loc_0065EB78: mov edx, eax
  loc_0065EB7A: lea ecx, var_AC
  loc_0065EB80: call [0040126Ch] ; __vbaStrMove
  loc_0065EB86: mov edx, eax
  loc_0065EB88: mov eax, var_68
  loc_0065EB8B: lea ecx, [eax+edi*4]
  loc_0065EB8E: call ebx
  loc_0065EB90: lea ecx, var_AC
  loc_0065EB96: call [0040129Ch] ; __vbaFreeStr
  loc_0065EB9C: lea ecx, var_C8
  loc_0065EBA2: call [00401020h] ; __vbaFreeVar
  loc_0065EBA8: lea ecx, var_1D8
  loc_0065EBAE: lea edx, var_1C8
  loc_0065EBB4: push ecx
  loc_0065EBB5: lea eax, var_20
  loc_0065EBB8: push edx
  loc_0065EBB9: push eax
  loc_0065EBBA: call [00401290h] ; __vbaVarForNext
  loc_0065EBC0: xor ebx, ebx
  loc_0065EBC2: mov edi, 00000002h
  loc_0065EBC7: jmp 0065E524h
  loc_0065EBCC: mov ecx, Me
  loc_0065EBCF: mov edx, 0042DFECh
  loc_0065EBD4: lea edi, [ecx+00000034h]
  loc_0065EBD7: mov ecx, edi
  loc_0065EBD9: call [004011FCh] ; __vbaStrCopy
  loc_0065EBDF: mov eax, 00000002h
  loc_0065EBE4: mov ecx, 00000001h
  loc_0065EBE9: mov var_138, eax
  loc_0065EBEF: mov var_148, eax
  loc_0065EBF5: mov var_158, eax
  loc_0065EBFB: lea edx, var_138
  loc_0065EC01: mov var_130, ecx
  loc_0065EC07: mov var_150, ecx
  loc_0065EC0D: lea eax, var_148
  loc_0065EC13: push edx
  loc_0065EC14: lea ecx, var_158
  loc_0065EC1A: push eax
  loc_0065EC1B: lea edx, var_1F8
  loc_0065EC21: push ecx
  loc_0065EC22: lea eax, var_1E8
  loc_0065EC28: push edx
  loc_0065EC29: lea ecx, var_20
  loc_0065EC2C: push eax
  loc_0065EC2D: push ecx
  loc_0065EC2E: mov var_140, 00000004h
  loc_0065EC38: call [004010A0h] ; __vbaVarForInit
  loc_0065EC3E: cmp eax, ebx
  loc_0065EC40: jz 0065ED9Dh
  loc_0065EC46: mov edx, [edi]
  loc_0065EC48: mov eax, var_68
  loc_0065EC4B: mov var_150, edx
  loc_0065EC51: lea ecx, var_C8
  loc_0065EC57: lea edx, var_20
  loc_0065EC5A: add eax, 00000004h
  loc_0065EC5D: push ecx
  loc_0065EC5E: push edx
  loc_0065EC5F: mov var_158, 00000008h
  loc_0065EC69: mov var_C0, 00000001h
  loc_0065EC73: mov var_C8, 00000002h
  loc_0065EC7D: mov var_130, eax
  loc_0065EC83: mov var_138, 00004008h
  loc_0065EC8D: call __vbaI4Var
  loc_0065EC8F: push eax
  loc_0065EC90: lea eax, var_138
  loc_0065EC96: lea ecx, var_D8
  loc_0065EC9C: push eax
  loc_0065EC9D: push ecx
  loc_0065EC9E: call [00401100h] ; rtcMidCharVar
  loc_0065ECA4: mov edx, var_68
  loc_0065ECA7: lea eax, var_F8
  loc_0065ECAD: lea ecx, var_20
  loc_0065ECB0: add edx, 00000008h
  loc_0065ECB3: push eax
  loc_0065ECB4: push ecx
  loc_0065ECB5: mov var_F0, 00000001h
  loc_0065ECBF: mov var_F8, 00000002h
  loc_0065ECC9: mov var_160, edx
  loc_0065ECCF: mov var_168, 00004008h
  loc_0065ECD9: call __vbaI4Var
  loc_0065ECDB: push eax
  loc_0065ECDC: lea edx, var_168
  loc_0065ECE2: lea eax, var_108
  loc_0065ECE8: push edx
  loc_0065ECE9: push eax
  loc_0065ECEA: call [00401100h] ; rtcMidCharVar
  loc_0065ECF0: lea ecx, var_158
  loc_0065ECF6: lea edx, var_D8
  loc_0065ECFC: push ecx
  loc_0065ECFD: lea eax, var_E8
  loc_0065ED03: push edx
  loc_0065ED04: push eax
  loc_0065ED05: call [004011C4h] ; __vbaVarCat
  loc_0065ED0B: lea ecx, var_108
  loc_0065ED11: push eax
  loc_0065ED12: lea edx, var_118
  loc_0065ED18: push ecx
  loc_0065ED19: push edx
  loc_0065ED1A: call [004011C4h] ; __vbaVarCat
  loc_0065ED20: push eax
  loc_0065ED21: call [00401034h] ; __vbaStrVarMove
  loc_0065ED27: mov edx, eax
  loc_0065ED29: lea ecx, var_AC
  loc_0065ED2F: call [0040126Ch] ; __vbaStrMove
  loc_0065ED35: mov edx, eax
  loc_0065ED37: mov ecx, edi
  loc_0065ED39: call [004011FCh] ; __vbaStrCopy
  loc_0065ED3F: lea ecx, var_AC
  loc_0065ED45: call [0040129Ch] ; __vbaFreeStr
  loc_0065ED4B: lea eax, var_118
  loc_0065ED51: lea ecx, var_108
  loc_0065ED57: push eax
  loc_0065ED58: lea edx, var_E8
  loc_0065ED5E: push ecx
  loc_0065ED5F: lea eax, var_F8
  loc_0065ED65: push edx
  loc_0065ED66: lea ecx, var_D8
  loc_0065ED6C: push eax
  loc_0065ED6D: lea edx, var_C8
  loc_0065ED73: push ecx
  loc_0065ED74: push edx
  loc_0065ED75: push 00000006h
  loc_0065ED77: call [0040103Ch] ; __vbaFreeVarList
  loc_0065ED7D: add esp, 0000001Ch
  loc_0065ED80: lea eax, var_1F8
  loc_0065ED86: lea ecx, var_1E8
  loc_0065ED8C: lea edx, var_20
  loc_0065ED8F: push eax
  loc_0065ED90: push ecx
  loc_0065ED91: push edx
  loc_0065ED92: call [00401290h] ; __vbaVarForNext
  loc_0065ED98: jmp 0065EC3Eh
  loc_0065ED9D: push 0065EEB9h
  loc_0065EDA2: jmp 0065EE0Dh
  loc_0065EDA4: lea eax, var_B0
  loc_0065EDAA: lea ecx, var_AC
  loc_0065EDB0: push eax
  loc_0065EDB1: push ecx
  loc_0065EDB2: push 00000002h
  loc_0065EDB4: call [00401204h] ; __vbaFreeStrList
  loc_0065EDBA: lea edx, var_B8
  loc_0065EDC0: lea eax, var_B4
  loc_0065EDC6: push edx
  loc_0065EDC7: push eax
  loc_0065EDC8: push 00000002h
  loc_0065EDCA: call [0040104Ch] ; __vbaFreeObjList
  loc_0065EDD0: lea ecx, var_128
  loc_0065EDD6: lea edx, var_118
  loc_0065EDDC: push ecx
  loc_0065EDDD: lea eax, var_108
  loc_0065EDE3: push edx
  loc_0065EDE4: lea ecx, var_F8
  loc_0065EDEA: push eax
  loc_0065EDEB: lea edx, var_E8
  loc_0065EDF1: push ecx
  loc_0065EDF2: lea eax, var_D8
  loc_0065EDF8: push edx
  loc_0065EDF9: lea ecx, var_C8
  loc_0065EDFF: push eax
  loc_0065EE00: push ecx
  loc_0065EE01: push 00000007h
  loc_0065EE03: call [0040103Ch] ; __vbaFreeVarList
  loc_0065EE09: add esp, 00000038h
  loc_0065EE0C: ret
  loc_0065EE0D: mov edi, [004012A0h] ; __vbaFreeObj
  loc_0065EE13: lea ecx, var_1B8
  loc_0065EE19: call edi
  loc_0065EE1B: lea edx, var_1F8
  loc_0065EE21: lea eax, var_1E8
  loc_0065EE27: push edx
  loc_0065EE28: lea ecx, var_1D8
  loc_0065EE2E: push eax
  loc_0065EE2F: lea edx, var_1C8
  loc_0065EE35: push ecx
  loc_0065EE36: push edx
  loc_0065EE37: push 00000004h
  loc_0065EE39: call [0040103Ch] ; __vbaFreeVarList
  loc_0065EE3F: add esp, 00000014h
  loc_0065EE42: lea ecx, var_20
  loc_0065EE45: call [00401020h] ; __vbaFreeVar
  loc_0065EE4B: mov esi, [00401090h] ; __vbaAryDestruct
  loc_0065EE51: lea ecx, var_18C
  loc_0065EE57: lea eax, var_38
  loc_0065EE5A: push ecx
  loc_0065EE5B: push 00000000h
  loc_0065EE5D: mov var_18C, eax
  loc_0065EE63: call __vbaAryDestruct
  loc_0065EE65: lea eax, var_190
  loc_0065EE6B: lea edx, var_58
  loc_0065EE6E: push eax
  loc_0065EE6F: push 00000000h
  loc_0065EE71: mov var_190, edx
  loc_0065EE77: call __vbaAryDestruct
  loc_0065EE79: lea edx, var_194
  loc_0065EE7F: lea ecx, var_74
  loc_0065EE82: push edx
  loc_0065EE83: push 00000000h
  loc_0065EE85: mov var_194, ecx
  loc_0065EE8B: call __vbaAryDestruct
  loc_0065EE8D: lea ecx, var_198
  loc_0065EE93: lea eax, var_90
  loc_0065EE99: push ecx
  loc_0065EE9A: push 00000000h
  loc_0065EE9C: mov var_198, eax
  loc_0065EEA2: call __vbaAryDestruct
  loc_0065EEA4: lea ecx, var_A4
  loc_0065EEAA: call [0040129Ch] ; __vbaFreeStr
  loc_0065EEB0: lea ecx, var_A8
  loc_0065EEB6: call edi
  loc_0065EEB8: ret
  loc_0065EEB9: mov ecx, var_10
  loc_0065EEBC: pop edi
  loc_0065EEBD: pop esi
  loc_0065EEBE: xor eax, eax
  loc_0065EEC0: mov fs:[00000000h], ecx
  loc_0065EEC7: pop ebx
  loc_0065EEC8: mov esp, ebp
  loc_0065EECA: pop ebp
  loc_0065EECB: retn 0004h
End Sub

Private Sub Proc_74_12_6603F0
  loc_006603F0: push ebp
  loc_006603F1: mov ebp, esp
  loc_006603F3: sub esp, 00000008h
  loc_006603F6: push 00405696h ; __vbaExceptHandler
  loc_006603FB: mov eax, fs:[00000000h]
  loc_00660401: push eax
  loc_00660402: mov fs:[00000000h], esp
  loc_00660409: sub esp, 000000C8h
  loc_0066040F: push ebx
  loc_00660410: push esi
  loc_00660411: push edi
  loc_00660412: mov var_8, esp
  loc_00660415: mov var_4, 004054C8h
  loc_0066041C: mov edi, [00401130h] ; __vbaAryConstruct2
  loc_00660422: push 00000002h
  loc_00660424: lea eax, var_2C
  loc_00660427: xor esi, esi
  loc_00660429: push 00438BD4h
  loc_0066042E: push eax
  loc_0066042F: mov var_50, esi
  loc_00660432: mov var_54, esi
  loc_00660435: mov var_58, esi
  loc_00660438: mov var_5C, esi
  loc_0066043B: mov var_60, esi
  loc_0066043E: mov var_70, esi
  loc_00660441: mov var_80, esi
  loc_00660444: mov var_90, esi
  loc_0066044A: mov var_A4, esi
  loc_00660450: mov var_A8, esi
  loc_00660456: call edi
  loc_00660458: push 00000008h
  loc_0066045A: lea ecx, var_48
  loc_0066045D: push 0043FA14h
  loc_00660462: push ecx
  loc_00660463: call edi
  loc_00660465: mov edi, Me
  loc_00660468: push edi
  loc_00660469: mov edx, [edi]
  loc_0066046B: call [edx+00000700h]
  loc_00660471: cmp eax, esi
  loc_00660473: jge 00660487h
  loc_00660475: push 00000700h
  loc_0066047A: push 00434D5Ch
  loc_0066047F: push edi
  loc_00660480: push eax
  loc_00660481: call [00401080h] ; __vbaHresultCheckObj
  loc_00660487: mov eax, [edi]
  loc_00660489: push edi
  loc_0066048A: call [eax+00000704h]
  loc_00660490: mov ecx, [edi+00000034h]
  loc_00660493: lea ebx, [edi+00000034h]
  loc_00660496: push 0043FD60h ; "&H"
  loc_0066049B: push ecx
  loc_0066049C: call [00401070h] ; __vbaStrCat
  loc_006604A2: mov esi, [0040126Ch] ; __vbaStrMove
  loc_006604A8: mov edx, eax
  loc_006604AA: lea ecx, var_50
  loc_006604AD: call __vbaStrMove
  loc_006604AF: push eax
  loc_006604B0: call [00401200h] ; __vbaI4Str
  loc_006604B6: push eax
  loc_006604B7: call [00401018h] ; __vbaStrI4
  loc_006604BD: mov edx, eax
  loc_006604BF: lea ecx, var_54
  loc_006604C2: call __vbaStrMove
  loc_006604C4: mov ecx, ebx
  loc_006604C6: mov ebx, [004011FCh] ; __vbaStrCopy
  loc_006604CC: mov edx, eax
  loc_006604CE: call ebx
  loc_006604D0: lea edx, var_54
  loc_006604D3: lea eax, var_50
  loc_006604D6: push edx
  loc_006604D7: push eax
  loc_006604D8: push 00000002h
  loc_006604DA: call [00401204h] ; __vbaFreeStrList
  loc_006604E0: add esp, 0000000Ch
  loc_006604E3: lea ecx, var_70
  loc_006604E6: mov var_68, 80020004h
  loc_006604ED: mov var_70, 0000000Ah
  loc_006604F4: push ecx
  loc_006604F5: call [004010A8h] ; rtcRandomize
  loc_006604FB: lea ecx, var_70
  loc_006604FE: call [00401020h] ; __vbaFreeVar
  loc_00660504: mov var_14, 00000001h
  loc_0066050B: mov eax, 00000002h
  loc_00660510: cmp var_14, ax
  loc_00660514: jg 00660621h
  loc_0066051A: lea edx, var_70
  loc_0066051D: mov var_68, 80020004h
  loc_00660524: push edx
  loc_00660525: mov var_70, 0000000Ah
  loc_0066052C: call [0040109Ch] ; rtcRandomNext
  loc_00660532: movsx edi, var_14
  loc_00660536: fstp real4 ptr var_A4
  loc_0066053C: cmp edi, 00000003h
  loc_0066053F: jb 00660547h
  loc_00660541: call [00401120h] ; __vbaGenerateBoundsError
  loc_00660547: fld real4 ptr var_A4
  loc_0066054D: fmul st0, real4 ptr [004054C4h]
  loc_00660553: fadd st0, real4 ptr [004054C0h]
  loc_00660559: fnstsw ax
  loc_0066055B: test al, 0Dh
  loc_0066055D: jnz 006608EFh
  loc_00660563: call [00401258h] ; __vbaR8IntI2
  loc_00660569: mov ecx, var_20
  loc_0066056C: mov [ecx+edi*2], ax
  loc_00660570: lea ecx, var_70
  loc_00660573: call [00401020h] ; __vbaFreeVar
  loc_00660579: cmp edi, 00000003h
  loc_0066057C: mov var_68, 00000008h
  loc_00660583: mov var_70, 00000002h
  loc_0066058A: jb 00660592h
  loc_0066058C: call [00401120h] ; __vbaGenerateBoundsError
  loc_00660592: mov edx, Me
  loc_00660595: mov ecx, var_20
  loc_00660598: add edx, 0000003Ch
  loc_0066059B: mov var_90, 00004008h
  loc_006605A5: mov var_88, edx
  loc_006605AB: lea eax, var_70
  loc_006605AE: movsx edx, [ecx+edi*2]
  loc_006605B2: push eax
  loc_006605B3: lea eax, var_90
  loc_006605B9: push edx
  loc_006605BA: lea ecx, var_80
  loc_006605BD: push eax
  loc_006605BE: push ecx
  loc_006605BF: call [00401100h] ; rtcMidCharVar
  loc_006605C5: cmp edi, 00000003h
  loc_006605C8: jb 006605D0h
  loc_006605CA: call [00401120h] ; __vbaGenerateBoundsError
  loc_006605D0: lea edx, var_80
  loc_006605D3: push edx
  loc_006605D4: call [00401034h] ; __vbaStrVarMove
  loc_006605DA: mov edx, eax
  loc_006605DC: lea ecx, var_50
  loc_006605DF: call __vbaStrMove
  loc_006605E1: mov edx, eax
  loc_006605E3: mov eax, var_3C
  loc_006605E6: lea ecx, [eax+edi*4]
  loc_006605E9: call ebx
  loc_006605EB: lea ecx, var_50
  loc_006605EE: call [0040129Ch] ; __vbaFreeStr
  loc_006605F4: lea ecx, var_80
  loc_006605F7: lea edx, var_70
  loc_006605FA: push ecx
  loc_006605FB: push edx
  loc_006605FC: push 00000002h
  loc_006605FE: call [0040103Ch] ; __vbaFreeVarList
  loc_00660604: mov edi, Me
  loc_00660607: mov eax, 00000001h
  loc_0066060C: add esp, 0000000Ch
  loc_0066060F: add ax, var_14
  loc_00660613: jo 006608F4h
  loc_00660619: mov var_14, eax
  loc_0066061C: jmp 0066050Bh
  loc_00660621: mov eax, var_3C
  loc_00660624: mov ecx, [eax+00000004h]
  loc_00660627: push ecx
  loc_00660628: call [004012A4h] ; rtcR8ValFromBstr
  loc_0066062E: mov edx, var_3C
  loc_00660631: fstp real8 ptr var_B0
  loc_00660637: mov eax, [edx+00000008h]
  loc_0066063A: push eax
  loc_0066063B: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660641: mov ecx, [edi+00000034h]
  loc_00660644: add edi, 00000034h
  loc_00660647: fstp real8 ptr var_B8
  loc_0066064D: push ecx
  loc_0066064E: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660654: fsub st0, real8 ptr var_B0
  loc_0066065A: lea edx, var_70
  loc_0066065D: mov var_70, 00000005h
  loc_00660664: push edx
  loc_00660665: fadd st0, real8 ptr var_B8
  loc_0066066B: fstp real8 ptr var_68
  loc_0066066E: fnstsw ax
  loc_00660670: test al, 0Dh
  loc_00660672: jnz 006608EFh
  loc_00660678: lea eax, var_80
  loc_0066067B: push eax
  loc_0066067C: call [00401240h] ; rtcVarStrFromVar
  loc_00660682: lea ecx, var_80
  loc_00660685: push ecx
  loc_00660686: call [00401034h] ; __vbaStrVarMove
  loc_0066068C: mov edx, eax
  loc_0066068E: lea ecx, var_50
  loc_00660691: call __vbaStrMove
  loc_00660693: mov edx, eax
  loc_00660695: mov ecx, edi
  loc_00660697: call ebx
  loc_00660699: lea ecx, var_50
  loc_0066069C: call [0040129Ch] ; __vbaFreeStr
  loc_006606A2: lea edx, var_80
  loc_006606A5: lea eax, var_70
  loc_006606A8: push edx
  loc_006606A9: push eax
  loc_006606AA: push 00000002h
  loc_006606AC: call [0040103Ch] ; __vbaFreeVarList
  loc_006606B2: mov edx, var_20
  loc_006606B5: mov ecx, [edi]
  loc_006606B7: add esp, 0000000Ch
  loc_006606BA: mov ax, [edx+00000002h]
  loc_006606BE: push ecx
  loc_006606BF: push eax
  loc_006606C0: call [0040100Ch] ; __vbaStrI2
  loc_006606C6: mov edx, eax
  loc_006606C8: lea ecx, var_50
  loc_006606CB: call __vbaStrMove
  loc_006606CD: push eax
  loc_006606CE: call [00401070h] ; __vbaStrCat
  loc_006606D4: mov edx, eax
  loc_006606D6: lea ecx, var_54
  loc_006606D9: call __vbaStrMove
  loc_006606DB: mov ecx, var_20
  loc_006606DE: push eax
  loc_006606DF: mov dx, [ecx+00000004h]
  loc_006606E3: push edx
  loc_006606E4: call [0040100Ch] ; __vbaStrI2
  loc_006606EA: mov edx, eax
  loc_006606EC: lea ecx, var_58
  loc_006606EF: call __vbaStrMove
  loc_006606F1: push eax
  loc_006606F2: call [00401070h] ; __vbaStrCat
  loc_006606F8: mov edx, eax
  loc_006606FA: lea ecx, var_5C
  loc_006606FD: call __vbaStrMove
  loc_006606FF: mov edx, eax
  loc_00660701: mov ecx, edi
  loc_00660703: call ebx
  loc_00660705: lea eax, var_5C
  loc_00660708: lea ecx, var_58
  loc_0066070B: push eax
  loc_0066070C: lea edx, var_54
  loc_0066070F: push ecx
  loc_00660710: push edx
  loc_00660711: lea eax, var_50
  loc_00660714: push eax
  loc_00660715: push 00000004h
  loc_00660717: call [00401204h] ; __vbaFreeStrList
  loc_0066071D: add esp, 00000014h
  loc_00660720: lea ecx, var_90
  loc_00660726: lea edx, var_70
  loc_00660729: mov var_88, edi
  loc_0066072F: push ecx
  loc_00660730: push edx
  loc_00660731: mov var_90, 00004008h
  loc_0066073B: call [004010DCh] ; rtcTrimVar
  loc_00660741: lea eax, var_70
  loc_00660744: push eax
  loc_00660745: call [00401034h] ; __vbaStrVarMove
  loc_0066074B: mov edx, eax
  loc_0066074D: lea ecx, var_50
  loc_00660750: call __vbaStrMove
  loc_00660752: mov edx, eax
  loc_00660754: mov ecx, edi
  loc_00660756: call ebx
  loc_00660758: lea ecx, var_50
  loc_0066075B: call [0040129Ch] ; __vbaFreeStr
  loc_00660761: lea ecx, var_70
  loc_00660764: call [00401020h] ; __vbaFreeVar
  loc_0066076A: mov eax, [0066A318h]
  loc_0066076F: test eax, eax
  loc_00660771: jnz 00660783h
  loc_00660773: push 0066A318h
  loc_00660778: push 0042E390h
  loc_0066077D: call [004011E8h] ; __vbaNew2
  loc_00660783: mov eax, [0066A318h]
  loc_00660788: lea edx, var_60
  loc_0066078B: push edx
  loc_0066078C: push eax
  loc_0066078D: mov ecx, [eax]
  loc_0066078F: mov var_BC, eax
  loc_00660795: call [ecx+00000014h]
  loc_00660798: test eax, eax
  loc_0066079A: fnclex
  loc_0066079C: jge 006607B3h
  loc_0066079E: mov ecx, var_BC
  loc_006607A4: push 00000014h
  loc_006607A6: push 0042E380h
  loc_006607AB: push ecx
  loc_006607AC: push eax
  loc_006607AD: call [00401080h] ; __vbaHresultCheckObj
  loc_006607B3: mov eax, var_60
  loc_006607B6: lea ecx, var_50
  loc_006607B9: push ecx
  loc_006607BA: push eax
  loc_006607BB: mov edx, [eax]
  loc_006607BD: mov var_C4, eax
  loc_006607C3: call [edx+00000050h]
  loc_006607C6: test eax, eax
  loc_006607C8: fnclex
  loc_006607CA: jge 006607E1h
  loc_006607CC: mov edx, var_C4
  loc_006607D2: push 00000050h
  loc_006607D4: push 0042E528h
  loc_006607D9: push edx
  loc_006607DA: push eax
  loc_006607DB: call [00401080h] ; __vbaHresultCheckObj
  loc_006607E1: mov eax, var_50
  loc_006607E4: push eax
  loc_006607E5: push 0043FA7Ch ; "\settings.ini"
  loc_006607EA: call [00401070h] ; __vbaStrCat
  loc_006607F0: mov edx, eax
  loc_006607F2: lea ecx, var_5C
  loc_006607F5: call __vbaStrMove
  loc_006607F7: mov edx, 0043FAF0h ; "MachineId"
  loc_006607FC: lea ecx, var_58
  loc_006607FF: call ebx
  loc_00660801: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_00660806: lea ecx, var_54
  loc_00660809: call ebx
  loc_0066080B: mov esi, Me
  loc_0066080E: lea edx, var_70
  loc_00660811: lea eax, var_5C
  loc_00660814: push edx
  loc_00660815: mov ecx, [esi]
  loc_00660817: push eax
  loc_00660818: lea edx, var_58
  loc_0066081B: push edi
  loc_0066081C: lea eax, var_54
  loc_0066081F: push edx
  loc_00660820: push eax
  loc_00660821: push esi
  loc_00660822: call [ecx+000006FCh]
  loc_00660828: test eax, eax
  loc_0066082A: jge 0066083Eh
  loc_0066082C: push 000006FCh
  loc_00660831: push 00434D5Ch
  loc_00660836: push esi
  loc_00660837: push eax
  loc_00660838: call [00401080h] ; __vbaHresultCheckObj
  loc_0066083E: lea ecx, var_5C
  loc_00660841: lea edx, var_58
  loc_00660844: push ecx
  loc_00660845: lea eax, var_54
  loc_00660848: push edx
  loc_00660849: lea ecx, var_50
  loc_0066084C: push eax
  loc_0066084D: push ecx
  loc_0066084E: push 00000004h
  loc_00660850: call [00401204h] ; __vbaFreeStrList
  loc_00660856: add esp, 00000014h
  loc_00660859: lea ecx, var_60
  loc_0066085C: call [004012A0h] ; __vbaFreeObj
  loc_00660862: lea ecx, var_70
  loc_00660865: call [00401020h] ; __vbaFreeVar
  loc_0066086B: fwait
  loc_0066086C: push 006608DAh
  loc_00660871: jmp 006608ABh
  loc_00660873: lea edx, var_5C
  loc_00660876: lea eax, var_58
  loc_00660879: push edx
  loc_0066087A: lea ecx, var_54
  loc_0066087D: push eax
  loc_0066087E: lea edx, var_50
  loc_00660881: push ecx
  loc_00660882: push edx
  loc_00660883: push 00000004h
  loc_00660885: call [00401204h] ; __vbaFreeStrList
  loc_0066088B: add esp, 00000014h
  loc_0066088E: lea ecx, var_60
  loc_00660891: call [004012A0h] ; __vbaFreeObj
  loc_00660897: lea eax, var_80
  loc_0066089A: lea ecx, var_70
  loc_0066089D: push eax
  loc_0066089E: push ecx
  loc_0066089F: push 00000002h
  loc_006608A1: call [0040103Ch] ; __vbaFreeVarList
  loc_006608A7: add esp, 0000000Ch
  loc_006608AA: ret
  loc_006608AB: mov esi, [00401090h] ; __vbaAryDestruct
  loc_006608B1: lea eax, var_A4
  loc_006608B7: lea edx, var_2C
  loc_006608BA: push eax
  loc_006608BB: push 00000000h
  loc_006608BD: mov var_A4, edx
  loc_006608C3: call __vbaAryDestruct
  loc_006608C5: lea edx, var_A8
  loc_006608CB: lea ecx, var_48
  loc_006608CE: push edx
  loc_006608CF: push 00000000h
  loc_006608D1: mov var_A8, ecx
  loc_006608D7: call __vbaAryDestruct
  loc_006608D9: ret
  loc_006608DA: mov ecx, var_10
  loc_006608DD: pop edi
  loc_006608DE: pop esi
  loc_006608DF: xor eax, eax
  loc_006608E1: mov fs:[00000000h], ecx
  loc_006608E8: pop ebx
  loc_006608E9: mov esp, ebp
  loc_006608EB: pop ebp
  loc_006608EC: retn 0004h
End Sub

Private Sub Proc_74_13_660900
  loc_00660900: push ebp
  loc_00660901: mov ebp, esp
  loc_00660903: sub esp, 00000008h
  loc_00660906: push 00405696h ; __vbaExceptHandler
  loc_0066090B: mov eax, fs:[00000000h]
  loc_00660911: push eax
  loc_00660912: mov fs:[00000000h], esp
  loc_00660919: sub esp, 000000D8h
  loc_0066091F: push ebx
  loc_00660920: push esi
  loc_00660921: push edi
  loc_00660922: mov var_8, esp
  loc_00660925: mov var_4, 004054D8h
  loc_0066092C: mov edi, [00401130h] ; __vbaAryConstruct2
  loc_00660932: push 00000002h
  loc_00660934: lea eax, var_34
  loc_00660937: xor esi, esi
  loc_00660939: push 00438BD4h
  loc_0066093E: push eax
  loc_0066093F: mov var_14, esi
  loc_00660942: mov var_1C, esi
  loc_00660945: mov var_58, esi
  loc_00660948: mov var_5C, esi
  loc_0066094B: mov var_60, esi
  loc_0066094E: mov var_64, esi
  loc_00660951: mov var_68, esi
  loc_00660954: mov var_6C, esi
  loc_00660957: mov var_7C, esi
  loc_0066095A: mov var_8C, esi
  loc_00660960: mov var_9C, esi
  loc_00660966: mov var_B0, esi
  loc_0066096C: mov var_B4, esi
  loc_00660972: call edi
  loc_00660974: push 00000008h
  loc_00660976: lea ecx, var_50
  loc_00660979: push 0043FA14h
  loc_0066097E: push ecx
  loc_0066097F: call edi
  loc_00660981: mov ebx, Me
  loc_00660984: push ebx
  loc_00660985: mov edx, [ebx]
  loc_00660987: call [edx+00000700h]
  loc_0066098D: cmp eax, esi
  loc_0066098F: jge 006609A3h
  loc_00660991: push 00000700h
  loc_00660996: push 00434D5Ch
  loc_0066099B: push ebx
  loc_0066099C: push eax
  loc_0066099D: call [00401080h] ; __vbaHresultCheckObj
  loc_006609A3: cmp [0066A318h], esi
  loc_006609A9: jnz 006609BBh
  loc_006609AB: push 0066A318h
  loc_006609B0: push 0042E390h
  loc_006609B5: call [004011E8h] ; __vbaNew2
  loc_006609BB: mov edi, [0066A318h]
  loc_006609C1: lea ecx, var_6C
  loc_006609C4: push ecx
  loc_006609C5: push edi
  loc_006609C6: mov eax, [edi]
  loc_006609C8: call [eax+00000014h]
  loc_006609CB: cmp eax, esi
  loc_006609CD: fnclex
  loc_006609CF: jge 006609E0h
  loc_006609D1: push 00000014h
  loc_006609D3: push 0042E380h
  loc_006609D8: push edi
  loc_006609D9: push eax
  loc_006609DA: call [00401080h] ; __vbaHresultCheckObj
  loc_006609E0: mov eax, var_6C
  loc_006609E3: lea ecx, var_58
  loc_006609E6: push ecx
  loc_006609E7: push eax
  loc_006609E8: mov edx, [eax]
  loc_006609EA: mov edi, eax
  loc_006609EC: call [edx+00000050h]
  loc_006609EF: cmp eax, esi
  loc_006609F1: fnclex
  loc_006609F3: jge 00660A04h
  loc_006609F5: push 00000050h
  loc_006609F7: push 0042E528h
  loc_006609FC: push edi
  loc_006609FD: push eax
  loc_006609FE: call [00401080h] ; __vbaHresultCheckObj
  loc_00660A04: mov edx, var_58
  loc_00660A07: push edx
  loc_00660A08: push 0043FA7Ch ; "\settings.ini"
  loc_00660A0D: call [00401070h] ; __vbaStrCat
  loc_00660A13: mov edx, eax
  loc_00660A15: lea ecx, var_64
  loc_00660A18: call [0040126Ch] ; __vbaStrMove
  loc_00660A1E: mov edi, [004011FCh] ; __vbaStrCopy
  loc_00660A24: mov edx, 0043E0BCh ; "RegistrationKey"
  loc_00660A29: lea ecx, var_60
  loc_00660A2C: call edi
  loc_00660A2E: mov edx, 0043EA0Ch ; "SETTINGS"
  loc_00660A33: lea ecx, var_5C
  loc_00660A36: call edi
  loc_00660A38: mov eax, [ebx]
  loc_00660A3A: lea ecx, var_68
  loc_00660A3D: lea edx, var_64
  loc_00660A40: push ecx
  loc_00660A41: push edx
  loc_00660A42: lea ecx, var_60
  loc_00660A45: lea edx, var_5C
  loc_00660A48: push ecx
  loc_00660A49: push edx
  loc_00660A4A: push ebx
  loc_00660A4B: call [eax+000006F8h]
  loc_00660A51: cmp eax, esi
  loc_00660A53: jge 00660A67h
  loc_00660A55: push 000006F8h
  loc_00660A5A: push 00434D5Ch
  loc_00660A5F: push ebx
  loc_00660A60: push eax
  loc_00660A61: call [00401080h] ; __vbaHresultCheckObj
  loc_00660A67: mov edx, var_68
  loc_00660A6A: lea ecx, var_1C
  loc_00660A6D: mov var_68, esi
  loc_00660A70: call [0040126Ch] ; __vbaStrMove
  loc_00660A76: lea eax, var_64
  loc_00660A79: lea ecx, var_60
  loc_00660A7C: push eax
  loc_00660A7D: lea edx, var_5C
  loc_00660A80: push ecx
  loc_00660A81: lea eax, var_58
  loc_00660A84: push edx
  loc_00660A85: push eax
  loc_00660A86: push 00000004h
  loc_00660A88: call [00401204h] ; __vbaFreeStrList
  loc_00660A8E: add esp, 00000014h
  loc_00660A91: lea ecx, var_6C
  loc_00660A94: call [004012A0h] ; __vbaFreeObj
  loc_00660A9A: mov ecx, var_1C
  loc_00660A9D: push ecx
  loc_00660A9E: push 0042DFECh
  loc_00660AA3: call [0040112Ch] ; __vbaStrCmp
  loc_00660AA9: test eax, eax
  loc_00660AAB: jnz 00660AF9h
  loc_00660AAD: mov edx, [ebx]
  loc_00660AAF: push ebx
  loc_00660AB0: call [edx+00000308h]
  loc_00660AB6: mov esi, [004010B8h] ; __vbaObjSet
  loc_00660ABC: push eax
  loc_00660ABD: lea eax, var_6C
  loc_00660AC0: push eax
  loc_00660AC1: call __vbaObjSet
  loc_00660AC3: mov edi, eax
  loc_00660AC5: push 00434FA8h ; "UNREGISTERED"
  loc_00660ACA: push edi
  loc_00660ACB: mov ecx, [edi]
  loc_00660ACD: call [ecx+000000A4h]
  loc_00660AD3: test eax, eax
  loc_00660AD5: fnclex
  loc_00660AD7: jge 00660AEBh
  loc_00660AD9: push 000000A4h
  loc_00660ADE: push 0042DFCCh
  loc_00660AE3: push edi
  loc_00660AE4: push eax
  loc_00660AE5: call [00401080h] ; __vbaHresultCheckObj
  loc_00660AEB: lea ecx, var_6C
  loc_00660AEE: call [004012A0h] ; __vbaFreeObj
  loc_00660AF4: jmp 00661088h
  loc_00660AF9: mov eax, 00000002h
  loc_00660AFE: mov edi, [00401100h] ; rtcMidCharVar
  loc_00660B04: mov var_74, eax
  loc_00660B07: mov var_7C, eax
  loc_00660B0A: lea edx, var_1C
  loc_00660B0D: lea eax, var_7C
  loc_00660B10: mov var_94, edx
  loc_00660B16: push eax
  loc_00660B17: lea ecx, var_9C
  loc_00660B1D: push 00000001h
  loc_00660B1F: lea edx, var_8C
  loc_00660B25: push ecx
  loc_00660B26: push edx
  loc_00660B27: mov var_9C, 00004008h
  loc_00660B31: call edi
  loc_00660B33: lea eax, var_8C
  loc_00660B39: push eax
  loc_00660B3A: call [004011D0h] ; __vbaI2Var
  loc_00660B40: mov ecx, var_28
  loc_00660B43: mov esi, [0040103Ch] ; __vbaFreeVarList
  loc_00660B49: lea edx, var_8C
  loc_00660B4F: mov [ecx+00000002h], ax
  loc_00660B53: lea eax, var_7C
  loc_00660B56: push edx
  loc_00660B57: push eax
  loc_00660B58: push 00000002h
  loc_00660B5A: call __vbaFreeVarList
  loc_00660B5C: add esp, 0000000Ch
  loc_00660B5F: mov eax, 00000002h
  loc_00660B64: lea ecx, var_1C
  loc_00660B67: lea edx, var_7C
  loc_00660B6A: mov var_74, eax
  loc_00660B6D: mov var_7C, eax
  loc_00660B70: mov var_94, ecx
  loc_00660B76: push edx
  loc_00660B77: lea eax, var_9C
  loc_00660B7D: push 00000003h
  loc_00660B7F: lea ecx, var_8C
  loc_00660B85: push eax
  loc_00660B86: push ecx
  loc_00660B87: mov var_9C, 00004008h
  loc_00660B91: call edi
  loc_00660B93: lea edx, var_8C
  loc_00660B99: push edx
  loc_00660B9A: call [004011D0h] ; __vbaI2Var
  loc_00660BA0: mov ecx, var_28
  loc_00660BA3: lea edx, var_8C
  loc_00660BA9: push edx
  loc_00660BAA: mov [ecx+00000004h], ax
  loc_00660BAE: lea eax, var_7C
  loc_00660BB1: push eax
  loc_00660BB2: push 00000002h
  loc_00660BB4: call __vbaFreeVarList
  loc_00660BB6: mov edx, var_1C
  loc_00660BB9: add esp, 0000000Ch
  loc_00660BBC: lea ecx, var_1C
  loc_00660BBF: mov var_9C, 00004008h
  loc_00660BC9: push edx
  loc_00660BCA: mov var_94, ecx
  loc_00660BD0: call [00401030h] ; __vbaLenBstr
  loc_00660BD6: sub eax, 00000004h
  loc_00660BD9: lea ecx, var_7C
  loc_00660BDC: jo 00661454h
  loc_00660BE2: push eax
  loc_00660BE3: lea eax, var_9C
  loc_00660BE9: push eax
  loc_00660BEA: push ecx
  loc_00660BEB: call [00401274h] ; rtcRightCharVar
  loc_00660BF1: lea edx, var_7C
  loc_00660BF4: push edx
  loc_00660BF5: call [00401034h] ; __vbaStrVarMove
  loc_00660BFB: mov edx, eax
  loc_00660BFD: lea ecx, var_1C
  loc_00660C00: call [0040126Ch] ; __vbaStrMove
  loc_00660C06: lea ecx, var_7C
  loc_00660C09: call [00401020h] ; __vbaFreeVar
  loc_00660C0F: lea eax, [ebx+0000003Ch]
  loc_00660C12: mov var_74, 0000000Bh
  loc_00660C19: mov var_7C, 00000002h
  loc_00660C20: mov var_94, eax
  loc_00660C26: mov var_9C, 00004008h
  loc_00660C30: mov ecx, var_28
  loc_00660C33: lea eax, var_7C
  loc_00660C36: push eax
  loc_00660C37: lea eax, var_9C
  loc_00660C3D: movsx edx, [ecx+00000002h]
  loc_00660C41: push edx
  loc_00660C42: lea ecx, var_8C
  loc_00660C48: push eax
  loc_00660C49: push ecx
  loc_00660C4A: call edi
  loc_00660C4C: lea edx, var_8C
  loc_00660C52: push edx
  loc_00660C53: call [00401034h] ; __vbaStrVarMove
  loc_00660C59: mov edx, eax
  loc_00660C5B: lea ecx, var_58
  loc_00660C5E: call [0040126Ch] ; __vbaStrMove
  loc_00660C64: mov edx, eax
  loc_00660C66: mov eax, var_44
  loc_00660C69: lea ecx, [eax+00000004h]
  loc_00660C6C: call [004011FCh] ; __vbaStrCopy
  loc_00660C72: lea ecx, var_58
  loc_00660C75: call [0040129Ch] ; __vbaFreeStr
  loc_00660C7B: lea ecx, var_8C
  loc_00660C81: lea edx, var_7C
  loc_00660C84: push ecx
  loc_00660C85: push edx
  loc_00660C86: push 00000002h
  loc_00660C88: call __vbaFreeVarList
  loc_00660C8A: mov ecx, var_28
  loc_00660C8D: lea eax, [ebx+0000003Ch]
  loc_00660C90: add esp, 0000000Ch
  loc_00660C93: mov var_74, 0000000Bh
  loc_00660C9A: mov var_7C, 00000002h
  loc_00660CA1: mov var_94, eax
  loc_00660CA7: mov var_9C, 00004008h
  loc_00660CB1: lea eax, var_7C
  loc_00660CB4: movsx edx, [ecx+00000004h]
  loc_00660CB8: push eax
  loc_00660CB9: lea eax, var_9C
  loc_00660CBF: push edx
  loc_00660CC0: lea ecx, var_8C
  loc_00660CC6: push eax
  loc_00660CC7: push ecx
  loc_00660CC8: call edi
  loc_00660CCA: lea edx, var_8C
  loc_00660CD0: push edx
  loc_00660CD1: call [00401034h] ; __vbaStrVarMove
  loc_00660CD7: mov edx, eax
  loc_00660CD9: lea ecx, var_58
  loc_00660CDC: call [0040126Ch] ; __vbaStrMove
  loc_00660CE2: mov edx, eax
  loc_00660CE4: mov eax, var_44
  loc_00660CE7: lea ecx, [eax+00000008h]
  loc_00660CEA: call [004011FCh] ; __vbaStrCopy
  loc_00660CF0: lea ecx, var_58
  loc_00660CF3: call [0040129Ch] ; __vbaFreeStr
  loc_00660CF9: lea ecx, var_8C
  loc_00660CFF: lea edx, var_7C
  loc_00660D02: push ecx
  loc_00660D03: push edx
  loc_00660D04: push 00000002h
  loc_00660D06: call __vbaFreeVarList
  loc_00660D08: mov eax, var_44
  loc_00660D0B: add esp, 0000000Ch
  loc_00660D0E: mov ecx, [eax+00000004h]
  loc_00660D11: push ecx
  loc_00660D12: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660D18: mov edx, var_44
  loc_00660D1B: fstp real8 ptr var_BC
  loc_00660D21: mov eax, [edx+00000008h]
  loc_00660D24: push eax
  loc_00660D25: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660D2B: mov ecx, var_1C
  loc_00660D2E: fstp real8 ptr var_C4
  loc_00660D34: push ecx
  loc_00660D35: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660D3B: fadd st0, real8 ptr var_BC
  loc_00660D41: fsub st0, real8 ptr var_C4
  loc_00660D47: fstp real8 ptr var_74
  loc_00660D4A: fnstsw ax
  loc_00660D4C: test al, 0Dh
  loc_00660D4E: jnz 0066144Fh
  loc_00660D54: lea edx, var_7C
  loc_00660D57: lea eax, var_8C
  loc_00660D5D: push edx
  loc_00660D5E: push eax
  loc_00660D5F: mov var_7C, 00000005h
  loc_00660D66: call [00401240h] ; rtcVarStrFromVar
  loc_00660D6C: lea ecx, var_8C
  loc_00660D72: push ecx
  loc_00660D73: call [00401034h] ; __vbaStrVarMove
  loc_00660D79: mov edx, eax
  loc_00660D7B: lea ecx, var_1C
  loc_00660D7E: call [0040126Ch] ; __vbaStrMove
  loc_00660D84: lea edx, var_8C
  loc_00660D8A: lea eax, var_7C
  loc_00660D8D: push edx
  loc_00660D8E: push eax
  loc_00660D8F: push 00000002h
  loc_00660D91: call __vbaFreeVarList
  loc_00660D93: mov eax, 00000002h
  loc_00660D98: add esp, 0000000Ch
  loc_00660D9B: mov var_74, eax
  loc_00660D9E: mov var_7C, eax
  loc_00660DA1: mov eax, var_1C
  loc_00660DA4: lea edx, var_7C
  loc_00660DA7: lea ecx, var_1C
  loc_00660DAA: push edx
  loc_00660DAB: push eax
  loc_00660DAC: mov var_94, ecx
  loc_00660DB2: mov var_9C, 00004008h
  loc_00660DBC: call [00401030h] ; __vbaLenBstr
  loc_00660DC2: sub eax, 00000003h
  loc_00660DC5: lea ecx, var_9C
  loc_00660DCB: jo 00661454h
  loc_00660DD1: push eax
  loc_00660DD2: lea edx, var_8C
  loc_00660DD8: push ecx
  loc_00660DD9: push edx
  loc_00660DDA: call edi
  loc_00660DDC: lea eax, var_8C
  loc_00660DE2: push eax
  loc_00660DE3: call [004011D0h] ; __vbaI2Var
  loc_00660DE9: mov ecx, var_28
  loc_00660DEC: lea edx, var_8C
  loc_00660DF2: push edx
  loc_00660DF3: mov [ecx+00000002h], ax
  loc_00660DF7: lea eax, var_7C
  loc_00660DFA: push eax
  loc_00660DFB: push 00000002h
  loc_00660DFD: call __vbaFreeVarList
  loc_00660DFF: mov eax, 00000002h
  loc_00660E04: add esp, 0000000Ch
  loc_00660E07: mov var_74, eax
  loc_00660E0A: mov var_7C, eax
  loc_00660E0D: mov eax, var_1C
  loc_00660E10: lea edx, var_7C
  loc_00660E13: lea ecx, var_1C
  loc_00660E16: push edx
  loc_00660E17: push eax
  loc_00660E18: mov var_94, ecx
  loc_00660E1E: mov var_9C, 00004008h
  loc_00660E28: call [00401030h] ; __vbaLenBstr
  loc_00660E2E: sub eax, 00000001h
  loc_00660E31: lea ecx, var_9C
  loc_00660E37: jo 00661454h
  loc_00660E3D: push eax
  loc_00660E3E: lea edx, var_8C
  loc_00660E44: push ecx
  loc_00660E45: push edx
  loc_00660E46: call edi
  loc_00660E48: lea eax, var_8C
  loc_00660E4E: push eax
  loc_00660E4F: call [004011D0h] ; __vbaI2Var
  loc_00660E55: mov ecx, var_28
  loc_00660E58: lea edx, var_8C
  loc_00660E5E: push edx
  loc_00660E5F: mov [ecx+00000004h], ax
  loc_00660E63: lea eax, var_7C
  loc_00660E66: push eax
  loc_00660E67: push 00000002h
  loc_00660E69: call __vbaFreeVarList
  loc_00660E6B: mov eax, 00000001h
  loc_00660E70: add esp, 0000000Ch
  loc_00660E73: mov var_18, eax
  loc_00660E76: mov ecx, 00000002h
  loc_00660E7B: cmp ax, cx
  loc_00660E7E: jg 00660F3Ch
  loc_00660E84: mov var_7C, ecx
  loc_00660E87: mov var_74, 00000008h
  loc_00660E8E: movsx ecx, ax
  loc_00660E91: cmp ecx, 00000003h
  loc_00660E94: jb 00660E9Fh
  loc_00660E96: call [00401120h] ; __vbaGenerateBoundsError
  loc_00660E9C: mov eax, var_18
  loc_00660E9F: mov edx, var_28
  loc_00660EA2: lea ecx, [ebx+0000003Ch]
  loc_00660EA5: movsx eax, ax
  loc_00660EA8: mov var_94, ecx
  loc_00660EAE: mov var_9C, 00004008h
  loc_00660EB8: movsx eax, [edx+eax*2]
  loc_00660EBC: lea ecx, var_7C
  loc_00660EBF: lea edx, var_8C
  loc_00660EC5: push ecx
  loc_00660EC6: lea ecx, var_9C
  loc_00660ECC: push eax
  loc_00660ECD: push ecx
  loc_00660ECE: push edx
  loc_00660ECF: call edi
  loc_00660ED1: movsx eax, var_18
  loc_00660ED5: cmp eax, 00000003h
  loc_00660ED8: jb 00660EE0h
  loc_00660EDA: call [00401120h] ; __vbaGenerateBoundsError
  loc_00660EE0: lea eax, var_8C
  loc_00660EE6: push eax
  loc_00660EE7: call [00401034h] ; __vbaStrVarMove
  loc_00660EED: mov edx, eax
  loc_00660EEF: lea ecx, var_58
  loc_00660EF2: call [0040126Ch] ; __vbaStrMove
  loc_00660EF8: mov ecx, var_44
  loc_00660EFB: mov edx, eax
  loc_00660EFD: movsx eax, var_18
  loc_00660F01: lea ecx, [ecx+eax*4]
  loc_00660F04: call [004011FCh] ; __vbaStrCopy
  loc_00660F0A: lea ecx, var_58
  loc_00660F0D: call [0040129Ch] ; __vbaFreeStr
  loc_00660F13: lea edx, var_8C
  loc_00660F19: lea eax, var_7C
  loc_00660F1C: push edx
  loc_00660F1D: push eax
  loc_00660F1E: push 00000002h
  loc_00660F20: call __vbaFreeVarList
  loc_00660F22: mov eax, 00000001h
  loc_00660F27: add esp, 0000000Ch
  loc_00660F2A: add ax, var_18
  loc_00660F2E: jo 00661454h
  loc_00660F34: mov var_18, eax
  loc_00660F37: jmp 00660E76h
  loc_00660F3C: mov ecx, var_1C
  loc_00660F3F: push ecx
  loc_00660F40: call [00401030h] ; __vbaLenBstr
  loc_00660F46: sub eax, 00000004h
  loc_00660F49: lea edx, var_1C
  loc_00660F4C: jo 00661454h
  loc_00660F52: mov var_74, eax
  loc_00660F55: lea eax, var_7C
  loc_00660F58: mov var_94, edx
  loc_00660F5E: push eax
  loc_00660F5F: lea ecx, var_9C
  loc_00660F65: push 00000001h
  loc_00660F67: lea edx, var_8C
  loc_00660F6D: push ecx
  loc_00660F6E: push edx
  loc_00660F6F: mov var_7C, 00000003h
  loc_00660F76: mov var_9C, 00004008h
  loc_00660F80: call edi
  loc_00660F82: lea eax, var_8C
  loc_00660F88: push eax
  loc_00660F89: call [00401034h] ; __vbaStrVarMove
  loc_00660F8F: mov edi, [0040126Ch] ; __vbaStrMove
  loc_00660F95: mov edx, eax
  loc_00660F97: lea ecx, var_1C
  loc_00660F9A: call edi
  loc_00660F9C: lea ecx, var_8C
  loc_00660FA2: lea edx, var_7C
  loc_00660FA5: push ecx
  loc_00660FA6: push edx
  loc_00660FA7: push 00000002h
  loc_00660FA9: call __vbaFreeVarList
  loc_00660FAB: mov eax, var_44
  loc_00660FAE: add esp, 0000000Ch
  loc_00660FB1: mov ecx, [eax+00000008h]
  loc_00660FB4: push ecx
  loc_00660FB5: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660FBB: mov edx, var_44
  loc_00660FBE: fstp real8 ptr var_BC
  loc_00660FC4: mov eax, [edx+00000004h]
  loc_00660FC7: push eax
  loc_00660FC8: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660FCE: mov ecx, var_1C
  loc_00660FD1: fstp real8 ptr var_C4
  loc_00660FD7: push ecx
  loc_00660FD8: call [004012A4h] ; rtcR8ValFromBstr
  loc_00660FDE: fsub st0, real8 ptr var_BC
  loc_00660FE4: sub esp, 00000008h
  loc_00660FE7: fadd st0, real8 ptr var_C4
  loc_00660FED: fnstsw ax
  loc_00660FEF: test al, 0Dh
  loc_00660FF1: jnz 0066144Fh
  loc_00660FF7: fstp real8 ptr [esp]
  loc_00660FFA: call [0040115Ch] ; __vbaStrR8
  loc_00661000: mov edx, eax
  loc_00661002: lea ecx, var_1C
  loc_00661005: call edi
  loc_00661007: lea eax, var_9C
  loc_0066100D: lea ecx, var_7C
  loc_00661010: lea edx, var_1C
  loc_00661013: push eax
  loc_00661014: push ecx
  loc_00661015: mov var_94, edx
  loc_0066101B: mov var_9C, 00004008h
  loc_00661025: call [004011F8h] ; rtcHexVarFromVar
  loc_0066102B: lea edx, var_7C
  loc_0066102E: push edx
  loc_0066102F: call [00401034h] ; __vbaStrVarMove
  loc_00661035: mov edx, eax
  loc_00661037: lea ecx, var_1C
  loc_0066103A: call edi
  loc_0066103C: lea ecx, var_7C
  loc_0066103F: call [00401020h] ; __vbaFreeVar
  loc_00661045: mov eax, var_1C
  loc_00661048: mov esi, [00401030h] ; __vbaLenBstr
  loc_0066104E: push eax
  loc_0066104F: call __vbaLenBstr
  loc_00661051: cmp eax, 00000008h
  loc_00661054: jge 00661079h
  loc_00661056: mov ecx, var_1C
  loc_00661059: push ecx
  loc_0066105A: call __vbaLenBstr
  loc_0066105C: cmp eax, 00000008h
  loc_0066105F: jz 00661079h
  loc_00661061: mov edx, var_1C
  loc_00661064: push 0042E3A4h
  loc_00661069: push edx
  loc_0066106A: call [00401070h] ; __vbaStrCat
  loc_00661070: mov edx, eax
  loc_00661072: lea ecx, var_1C
  loc_00661075: call edi
  loc_00661077: jmp 00661056h
  loc_00661079: mov eax, [ebx]
  loc_0066107B: push ebx
  loc_0066107C: call [eax+00000704h]
  loc_00661082: mov esi, [004010B8h] ; __vbaObjSet
  loc_00661088: mov ecx, var_1C
  loc_0066108B: mov edx, [ebx+00000034h]
  loc_0066108E: push ecx
  loc_0066108F: push edx
  loc_00661090: call [0040112Ch] ; __vbaStrCmp
  loc_00661096: test eax, eax
  loc_00661098: jnz 006610A2h
  loc_0066109A: mov [ebx+0000003Ah], FFFFFFh
  loc_006610A0: jmp 006610A8h
  loc_006610A2: mov [ebx+0000003Ah], 0000h
  loc_006610A8: cmp [ebx+0000003Ah], FFFFFFh
  loc_006610AD: mov eax, [ebx]
  loc_006610AF: push ebx
  loc_006610B0: jnz 006611D7h
  loc_006610B6: call [eax+00000300h]
  loc_006610BC: lea ecx, var_6C
  loc_006610BF: push eax
  loc_006610C0: push ecx
  loc_006610C1: call __vbaObjSet
  loc_006610C3: mov edi, eax
  loc_006610C5: push 0043FD6Ch ; "Unregister"
  loc_006610CA: push edi
  loc_006610CB: mov edx, [edi]
  loc_006610CD: call [edx+00000054h]
  loc_006610D0: test eax, eax
  loc_006610D2: fnclex
  loc_006610D4: jge 006610E5h
  loc_006610D6: push 00000054h
  loc_006610D8: push 0043FC70h
  loc_006610DD: push edi
  loc_006610DE: push eax
  loc_006610DF: call [00401080h] ; __vbaHresultCheckObj
  loc_006610E5: lea ecx, var_6C
  loc_006610E8: call [004012A0h] ; __vbaFreeObj
  loc_006610EE: mov eax, [ebx]
  loc_006610F0: push ebx
  loc_006610F1: call [eax+00000308h]
  loc_006610F7: lea ecx, var_6C
  loc_006610FA: push eax
  loc_006610FB: push ecx
  loc_006610FC: call __vbaObjSet
  loc_006610FE: mov edi, eax
  loc_00661100: push 0043FD88h ; "REGISTERED"
  loc_00661105: push edi
  loc_00661106: mov edx, [edi]
  loc_00661108: call [edx+000000A4h]
  loc_0066110E: test eax, eax
  loc_00661110: fnclex
  loc_00661112: jge 00661126h
  loc_00661114: push 000000A4h
  loc_00661119: push 0042DFCCh
  loc_0066111E: push edi
  loc_0066111F: push eax
  loc_00661120: call [00401080h] ; __vbaHresultCheckObj
  loc_00661126: lea ecx, var_6C
  loc_00661129: call [004012A0h] ; __vbaFreeObj
  loc_0066112F: mov eax, [ebx]
  loc_00661131: push ebx
  loc_00661132: call [eax+00000334h]
  loc_00661138: lea ecx, var_6C
  loc_0066113B: push eax
  loc_0066113C: push ecx
  loc_0066113D: call __vbaObjSet
  loc_0066113F: mov edi, eax
  loc_00661141: push 00000000h
  loc_00661143: push edi
  loc_00661144: mov edx, [edi]
  loc_00661146: call [edx+0000008Ch]
  loc_0066114C: test eax, eax
  loc_0066114E: fnclex
  loc_00661150: jge 00661164h
  loc_00661152: push 0000008Ch
  loc_00661157: push 0043FC70h
  loc_0066115C: push edi
  loc_0066115D: push eax
  loc_0066115E: call [00401080h] ; __vbaHresultCheckObj
  loc_00661164: lea ecx, var_6C
  loc_00661167: call [004012A0h] ; __vbaFreeObj
  loc_0066116D: mov eax, [ebx]
  loc_0066116F: push ebx
  loc_00661170: call [eax+00000318h]
  loc_00661176: lea ecx, var_6C
  loc_00661179: push eax
  loc_0066117A: push ecx
  loc_0066117B: call __vbaObjSet
  loc_0066117D: mov edi, eax
  loc_0066117F: push 00000000h
  loc_00661181: push edi
  loc_00661182: mov edx, [edi]
  loc_00661184: call [edx+0000008Ch]
  loc_0066118A: test eax, eax
  loc_0066118C: fnclex
  loc_0066118E: jge 006611A2h
  loc_00661190: push 0000008Ch
  loc_00661195: push 0043FC70h
  loc_0066119A: push edi
  loc_0066119B: push eax
  loc_0066119C: call [00401080h] ; __vbaHresultCheckObj
  loc_006611A2: lea ecx, var_6C
  loc_006611A5: call [004012A0h] ; __vbaFreeObj
  loc_006611AB: mov eax, [ebx]
  loc_006611AD: push ebx
  loc_006611AE: call [eax+00000300h]
  loc_006611B4: lea ecx, var_6C
  loc_006611B7: push eax
  loc_006611B8: push ecx
  loc_006611B9: call __vbaObjSet
  loc_006611BB: mov edi, eax
  loc_006611BD: push FFFFFFFFh
  loc_006611BF: push edi
  loc_006611C0: mov edx, [edi]
  loc_006611C2: call [edx+0000008Ch]
  loc_006611C8: test eax, eax
  loc_006611CA: fnclex
  loc_006611CC: jge 00661285h
  loc_006611D2: jmp 00661273h
  loc_006611D7: call [eax+00000300h]
  loc_006611DD: lea ecx, var_6C
  loc_006611E0: push eax
  loc_006611E1: push ecx
  loc_006611E2: call __vbaObjSet
  loc_006611E4: mov edi, eax
  loc_006611E6: push 0043FDA4h ; "Register"
  loc_006611EB: push edi
  loc_006611EC: mov edx, [edi]
  loc_006611EE: call [edx+00000054h]
  loc_006611F1: test eax, eax
  loc_006611F3: fnclex
  loc_006611F5: jge 00661206h
  loc_006611F7: push 00000054h
  loc_006611F9: push 0043FC70h
  loc_006611FE: push edi
  loc_006611FF: push eax
  loc_00661200: call [00401080h] ; __vbaHresultCheckObj
  loc_00661206: lea ecx, var_6C
  loc_00661209: call [004012A0h] ; __vbaFreeObj
  loc_0066120F: mov eax, [ebx]
  loc_00661211: push ebx
  loc_00661212: call [eax+00000308h]
  loc_00661218: lea ecx, var_6C
  loc_0066121B: push eax
  loc_0066121C: push ecx
  loc_0066121D: call __vbaObjSet
  loc_0066121F: mov edi, eax
  loc_00661221: push 00434FA8h ; "UNREGISTERED"
  loc_00661226: push edi
  loc_00661227: mov edx, [edi]
  loc_00661229: call [edx+000000A4h]
  loc_0066122F: test eax, eax
  loc_00661231: fnclex
  loc_00661233: jge 00661247h
  loc_00661235: push 000000A4h
  loc_0066123A: push 0042DFCCh
  loc_0066123F: push edi
  loc_00661240: push eax
  loc_00661241: call [00401080h] ; __vbaHresultCheckObj
  loc_00661247: lea ecx, var_6C
  loc_0066124A: call [004012A0h] ; __vbaFreeObj
  loc_00661250: mov eax, [ebx]
  loc_00661252: push ebx
  loc_00661253: call [eax+00000334h]
  loc_00661259: lea ecx, var_6C
  loc_0066125C: push eax
  loc_0066125D: push ecx
  loc_0066125E: call __vbaObjSet
  loc_00661260: mov edi, eax
  loc_00661262: push FFFFFFFFh
  loc_00661264: push edi
  loc_00661265: mov edx, [edi]
  loc_00661267: call [edx+0000008Ch]
  loc_0066126D: test eax, eax
  loc_0066126F: fnclex
  loc_00661271: jge 00661285h
  loc_00661273: push 0000008Ch
  loc_00661278: push 0043FC70h
  loc_0066127D: push edi
  loc_0066127E: push eax
  loc_0066127F: call [00401080h] ; __vbaHresultCheckObj
  loc_00661285: lea ecx, var_6C
  loc_00661288: call [004012A0h] ; __vbaFreeObj
  loc_0066128E: mov eax, [ebx]
  loc_00661290: push ebx
  loc_00661291: call [eax+00000308h]
  loc_00661297: lea ecx, var_6C
  loc_0066129A: push eax
  loc_0066129B: push ecx
  loc_0066129C: call __vbaObjSet
  loc_0066129E: mov edi, eax
  loc_006612A0: push edi
  loc_006612A1: mov edx, [edi]
  loc_006612A3: call [edx+0000021Ch]
  loc_006612A9: test eax, eax
  loc_006612AB: fnclex
  loc_006612AD: jge 006612C1h
  loc_006612AF: push 0000021Ch
  loc_006612B4: push 0042DFCCh
  loc_006612B9: push edi
  loc_006612BA: push eax
  loc_006612BB: call [00401080h] ; __vbaHresultCheckObj
  loc_006612C1: lea ecx, var_6C
  loc_006612C4: call [004012A0h] ; __vbaFreeObj
  loc_006612CA: mov eax, [006682A0h]
  loc_006612CF: test eax, eax
  loc_006612D1: jnz 006612E3h
  loc_006612D3: push 006682A0h
  loc_006612D8: push 00414A94h
  loc_006612DD: call [004011E8h] ; __vbaNew2
  loc_006612E3: mov eax, [ebx]
  loc_006612E5: mov edi, [006682A0h]
  loc_006612EB: push ebx
  loc_006612EC: call [eax+00000308h]
  loc_006612F2: lea ecx, var_6C
  loc_006612F5: push eax
  loc_006612F6: push ecx
  loc_006612F7: call __vbaObjSet
  loc_006612F9: mov esi, eax
  loc_006612FB: lea eax, var_58
  loc_006612FE: push eax
  loc_006612FF: push esi
  loc_00661300: mov edx, [esi]
  loc_00661302: call [edx+000000A0h]
  loc_00661308: test eax, eax
  loc_0066130A: fnclex
  loc_0066130C: jge 00661324h
  loc_0066130E: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_00661314: push 000000A0h
  loc_00661319: push 0042DFCCh
  loc_0066131E: push esi
  loc_0066131F: push eax
  loc_00661320: call ebx
  loc_00661322: jmp 0066132Ah
  loc_00661324: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0066132A: mov ecx, var_58
  loc_0066132D: mov esi, [edi]
  loc_0066132F: push 0043FDBCh ; "Registration Demo - "
  loc_00661334: push ecx
  loc_00661335: call [00401070h] ; __vbaStrCat
  loc_0066133B: mov edx, eax
  loc_0066133D: lea ecx, var_5C
  loc_00661340: call [0040126Ch] ; __vbaStrMove
  loc_00661346: push eax
  loc_00661347: push edi
  loc_00661348: call [esi+00000054h]
  loc_0066134B: test eax, eax
  loc_0066134D: fnclex
  loc_0066134F: jge 0066135Ch
  loc_00661351: push 00000054h
  loc_00661353: push 00434D2Ch
  loc_00661358: push edi
  loc_00661359: push eax
  loc_0066135A: call ebx
  loc_0066135C: lea edx, var_5C
  loc_0066135F: lea eax, var_58
  loc_00661362: push edx
  loc_00661363: push eax
  loc_00661364: push 00000002h
  loc_00661366: call [00401204h] ; __vbaFreeStrList
  loc_0066136C: add esp, 0000000Ch
  loc_0066136F: lea ecx, var_6C
  loc_00661372: call [004012A0h] ; __vbaFreeObj
  loc_00661378: mov eax, [006682A0h]
  loc_0066137D: test eax, eax
  loc_0066137F: jnz 00661391h
  loc_00661381: push 006682A0h
  loc_00661386: push 00414A94h
  loc_0066138B: call [004011E8h] ; __vbaNew2
  loc_00661391: mov esi, [006682A0h]
  loc_00661397: push esi
  loc_00661398: mov ecx, [esi]
  loc_0066139A: call [ecx+000002A0h]
  loc_006613A0: test eax, eax
  loc_006613A2: fnclex
  loc_006613A4: jge 006613B4h
  loc_006613A6: push 000002A0h
  loc_006613AB: push 00434D2Ch
  loc_006613B0: push esi
  loc_006613B1: push eax
  loc_006613B2: call ebx
  loc_006613B4: fwait
  loc_006613B5: push 0066143Ah
  loc_006613BA: jmp 006613FBh
  loc_006613BC: lea edx, var_68
  loc_006613BF: lea eax, var_64
  loc_006613C2: push edx
  loc_006613C3: lea ecx, var_60
  loc_006613C6: push eax
  loc_006613C7: lea edx, var_5C
  loc_006613CA: push ecx
  loc_006613CB: lea eax, var_58
  loc_006613CE: push edx
  loc_006613CF: push eax
  loc_006613D0: push 00000005h
  loc_006613D2: call [00401204h] ; __vbaFreeStrList
  loc_006613D8: add esp, 00000018h
  loc_006613DB: lea ecx, var_6C
  loc_006613DE: call [004012A0h] ; __vbaFreeObj
  loc_006613E4: lea ecx, var_8C
  loc_006613EA: lea edx, var_7C
  loc_006613ED: push ecx
  loc_006613EE: push edx
  loc_006613EF: push 00000002h
  loc_006613F1: call [0040103Ch] ; __vbaFreeVarList
  loc_006613F7: add esp, 0000000Ch
  loc_006613FA: ret
  loc_006613FB: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_00661401: lea ecx, var_14
  loc_00661404: call __vbaFreeStr
  loc_00661406: lea ecx, var_1C
  loc_00661409: call __vbaFreeStr
  loc_0066140B: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00661411: lea ecx, var_B0
  loc_00661417: lea eax, var_34
  loc_0066141A: push ecx
  loc_0066141B: push 00000000h
  loc_0066141D: mov var_B0, eax
  loc_00661423: call __vbaAryDestruct
  loc_00661425: lea eax, var_B4
  loc_0066142B: lea edx, var_50
  loc_0066142E: push eax
  loc_0066142F: push 00000000h
  loc_00661431: mov var_B4, edx
  loc_00661437: call __vbaAryDestruct
  loc_00661439: ret
  loc_0066143A: mov ecx, var_10
  loc_0066143D: pop edi
  loc_0066143E: pop esi
  loc_0066143F: xor eax, eax
  loc_00661441: mov fs:[00000000h], ecx
  loc_00661448: pop ebx
  loc_00661449: mov esp, ebp
  loc_0066144B: pop ebp
  loc_0066144C: retn 0004h
End Sub

Private Sub Proc_74_14_661F20
  loc_00661F20: push ebp
  loc_00661F21: mov ebp, esp
  loc_00661F23: sub esp, 00000008h
  loc_00661F26: push 00405696h ; __vbaExceptHandler
  loc_00661F2B: mov eax, fs:[00000000h]
  loc_00661F31: push eax
  loc_00661F32: mov fs:[00000000h], esp
  loc_00661F39: sub esp, 0000012Ch
  loc_00661F3F: push ebx
  loc_00661F40: push esi
  loc_00661F41: push edi
  loc_00661F42: mov var_8, esp
  loc_00661F45: mov var_4, 00405528h
  loc_00661F4C: mov ebx, [00401130h] ; __vbaAryConstruct2
  loc_00661F52: mov esi, 00000002h
  loc_00661F57: push esi
  loc_00661F58: lea eax, var_34
  loc_00661F5B: xor edi, edi
  loc_00661F5D: push 00438BD4h
  loc_00661F62: push eax
  loc_00661F63: mov var_14, edi
  loc_00661F66: mov var_58, edi
  loc_00661F69: mov var_5C, edi
  loc_00661F6C: mov var_6C, edi
  loc_00661F6F: mov var_7C, edi
  loc_00661F72: mov var_8C, edi
  loc_00661F78: mov var_9C, edi
  loc_00661F7E: mov var_AC, edi
  loc_00661F84: mov var_BC, edi
  loc_00661F8A: mov var_CC, edi
  loc_00661F90: mov var_DC, edi
  loc_00661F96: mov var_110, edi
  loc_00661F9C: mov var_114, edi
  loc_00661FA2: call ebx
  loc_00661FA4: push 00000008h
  loc_00661FA6: lea ecx, var_50
  loc_00661FA9: push 0043FA14h
  loc_00661FAE: push ecx
  loc_00661FAF: call ebx
  loc_00661FB1: mov ebx, Me
  loc_00661FB4: push ebx
  loc_00661FB5: mov edx, [ebx]
  loc_00661FB7: call [edx+00000700h]
  loc_00661FBD: cmp eax, edi
  loc_00661FBF: jge 00661FD3h
  loc_00661FC1: push 00000700h
  loc_00661FC6: push 00434D5Ch
  loc_00661FCB: push ebx
  loc_00661FCC: push eax
  loc_00661FCD: call [00401080h] ; __vbaHresultCheckObj
  loc_00661FD3: mov eax, [ebx]
  loc_00661FD5: push ebx
  loc_00661FD6: call [eax+00000330h]
  loc_00661FDC: lea ecx, var_5C
  loc_00661FDF: push eax
  loc_00661FE0: push ecx
  loc_00661FE1: call [004010B8h] ; __vbaObjSet
  loc_00661FE7: mov ebx, eax
  loc_00661FE9: lea eax, var_58
  loc_00661FEC: push eax
  loc_00661FED: push ebx
  loc_00661FEE: mov edx, [ebx]
  loc_00661FF0: call [edx+000000A0h]
  loc_00661FF6: cmp eax, edi
  loc_00661FF8: fnclex
  loc_00661FFA: jge 0066200Eh
  loc_00661FFC: push 000000A0h
  loc_00662001: push 0042DFCCh
  loc_00662006: push ebx
  loc_00662007: push eax
  loc_00662008: call [00401080h] ; __vbaHresultCheckObj
  loc_0066200E: mov edx, var_58
  loc_00662011: mov ebx, [0040126Ch] ; __vbaStrMove
  loc_00662017: lea ecx, var_14
  loc_0066201A: mov var_58, edi
  loc_0066201D: call ebx
  loc_0066201F: lea ecx, var_5C
  loc_00662022: call [004012A0h] ; __vbaFreeObj
  loc_00662028: mov eax, var_14
  loc_0066202B: mov edi, [00401030h] ; __vbaLenBstr
  loc_00662031: lea edx, var_6C
  loc_00662034: lea ecx, var_14
  loc_00662037: push edx
  loc_00662038: push eax
  loc_00662039: mov var_64, esi
  loc_0066203C: mov var_6C, esi
  loc_0066203F: mov var_C4, ecx
  loc_00662045: mov var_CC, 00004008h
  loc_0066204F: call edi
  loc_00662051: sub eax, 00000003h
  loc_00662054: lea ecx, var_CC
  loc_0066205A: jo 00662489h
  loc_00662060: push eax
  loc_00662061: lea edx, var_7C
  loc_00662064: push ecx
  loc_00662065: push edx
  loc_00662066: call [00401100h] ; rtcMidCharVar
  loc_0066206C: mov esi, [004011D0h] ; __vbaI2Var
  loc_00662072: lea eax, var_7C
  loc_00662075: push eax
  loc_00662076: call __vbaI2Var
  loc_00662078: mov ecx, var_28
  loc_0066207B: lea edx, var_7C
  loc_0066207E: push edx
  loc_0066207F: mov [ecx+00000002h], ax
  loc_00662083: lea eax, var_6C
  loc_00662086: push eax
  loc_00662087: push 00000002h
  loc_00662089: call [0040103Ch] ; __vbaFreeVarList
  loc_0066208F: mov eax, 00000002h
  loc_00662094: add esp, 0000000Ch
  loc_00662097: mov var_64, eax
  loc_0066209A: mov var_6C, eax
  loc_0066209D: mov eax, var_14
  loc_006620A0: lea edx, var_6C
  loc_006620A3: lea ecx, var_14
  loc_006620A6: push edx
  loc_006620A7: push eax
  loc_006620A8: mov var_C4, ecx
  loc_006620AE: mov var_CC, 00004008h
  loc_006620B8: call edi
  loc_006620BA: sub eax, 00000001h
  loc_006620BD: lea ecx, var_CC
  loc_006620C3: jo 00662489h
  loc_006620C9: push eax
  loc_006620CA: lea edx, var_7C
  loc_006620CD: push ecx
  loc_006620CE: push edx
  loc_006620CF: call [00401100h] ; rtcMidCharVar
  loc_006620D5: lea eax, var_7C
  loc_006620D8: push eax
  loc_006620D9: call __vbaI2Var
  loc_006620DB: mov ecx, var_28
  loc_006620DE: lea edx, var_7C
  loc_006620E1: push edx
  loc_006620E2: mov [ecx+00000004h], ax
  loc_006620E6: lea eax, var_6C
  loc_006620E9: push eax
  loc_006620EA: push 00000002h
  loc_006620EC: call [0040103Ch] ; __vbaFreeVarList
  loc_006620F2: add esp, 0000000Ch
  loc_006620F5: mov esi, 00000001h
  loc_006620FA: mov eax, 00000002h
  loc_006620FF: cmp si, ax
  loc_00662102: jg 006621B4h
  loc_00662108: movsx edi, si
  loc_0066210B: cmp edi, 00000003h
  loc_0066210E: mov var_64, 00000008h
  loc_00662115: mov var_6C, eax
  loc_00662118: jb 00662120h
  loc_0066211A: call [00401120h] ; __vbaGenerateBoundsError
  loc_00662120: mov ecx, Me
  loc_00662123: mov eax, var_28
  loc_00662126: add ecx, 0000003Ch
  loc_00662129: mov var_CC, 00004008h
  loc_00662133: mov var_C4, ecx
  loc_00662139: lea edx, var_6C
  loc_0066213C: movsx ecx, [eax+edi*2]
  loc_00662140: push edx
  loc_00662141: lea edx, var_CC
  loc_00662147: push ecx
  loc_00662148: lea eax, var_7C
  loc_0066214B: push edx
  loc_0066214C: push eax
  loc_0066214D: call [00401100h] ; rtcMidCharVar
  loc_00662153: cmp edi, 00000003h
  loc_00662156: jb 0066215Eh
  loc_00662158: call [00401120h] ; __vbaGenerateBoundsError
  loc_0066215E: lea ecx, var_7C
  loc_00662161: push ecx
  loc_00662162: call [00401034h] ; __vbaStrVarMove
  loc_00662168: mov edx, eax
  loc_0066216A: lea ecx, var_58
  loc_0066216D: call ebx
  loc_0066216F: mov edx, eax
  loc_00662171: mov eax, var_44
  loc_00662174: lea ecx, [eax+edi*4]
  loc_00662177: call [004011FCh] ; __vbaStrCopy
  loc_0066217D: lea ecx, var_58
  loc_00662180: call [0040129Ch] ; __vbaFreeStr
  loc_00662186: lea ecx, var_7C
  loc_00662189: lea edx, var_6C
  loc_0066218C: push ecx
  loc_0066218D: push edx
  loc_0066218E: push 00000002h
  loc_00662190: call [0040103Ch] ; __vbaFreeVarList
  loc_00662196: mov edi, [00401030h] ; __vbaLenBstr
  loc_0066219C: mov eax, 00000001h
  loc_006621A1: add esp, 0000000Ch
  loc_006621A4: add ax, si
  loc_006621A7: jo 00662489h
  loc_006621AD: mov esi, eax
  loc_006621AF: jmp 006620FAh
  loc_006621B4: mov eax, var_14
  loc_006621B7: push eax
  loc_006621B8: call edi
  loc_006621BA: sub eax, 00000004h
  loc_006621BD: lea ecx, var_14
  loc_006621C0: jo 00662489h
  loc_006621C6: lea edx, var_6C
  loc_006621C9: mov var_64, eax
  loc_006621CC: mov var_C4, ecx
  loc_006621D2: push edx
  loc_006621D3: lea eax, var_CC
  loc_006621D9: push 00000001h
  loc_006621DB: lea ecx, var_7C
  loc_006621DE: push eax
  loc_006621DF: push ecx
  loc_006621E0: mov var_6C, 00000003h
  loc_006621E7: mov var_CC, 00004008h
  loc_006621F1: call [00401100h] ; rtcMidCharVar
  loc_006621F7: mov esi, [00401034h] ; __vbaStrVarMove
  loc_006621FD: lea edx, var_7C
  loc_00662200: push edx
  loc_00662201: call __vbaStrVarMove
  loc_00662203: mov edx, eax
  loc_00662205: lea ecx, var_14
  loc_00662208: call ebx
  loc_0066220A: lea eax, var_7C
  loc_0066220D: lea ecx, var_6C
  loc_00662210: push eax
  loc_00662211: push ecx
  loc_00662212: push 00000002h
  loc_00662214: call [0040103Ch] ; __vbaFreeVarList
  loc_0066221A: mov edx, var_44
  loc_0066221D: add esp, 0000000Ch
  loc_00662220: mov eax, [edx+00000008h]
  loc_00662223: push eax
  loc_00662224: call [004012A4h] ; rtcR8ValFromBstr
  loc_0066222A: mov ecx, var_44
  loc_0066222D: fstp real8 ptr var_11C
  loc_00662233: mov edx, [ecx+00000004h]
  loc_00662236: push edx
  loc_00662237: call [004012A4h] ; rtcR8ValFromBstr
  loc_0066223D: mov eax, var_14
  loc_00662240: fstp real8 ptr var_124
  loc_00662246: push eax
  loc_00662247: call [004012A4h] ; rtcR8ValFromBstr
  loc_0066224D: fsub st0, real8 ptr var_11C
  loc_00662253: sub esp, 00000008h
  loc_00662256: fadd st0, real8 ptr var_124
  loc_0066225C: fnstsw ax
  loc_0066225E: test al, 0Dh
  loc_00662260: jnz 00662484h
  loc_00662266: fstp real8 ptr [esp]
  loc_00662269: call [0040115Ch] ; __vbaStrR8
  loc_0066226F: mov edx, eax
  loc_00662271: lea ecx, var_14
  loc_00662274: call ebx
  loc_00662276: lea edx, var_CC
  loc_0066227C: lea eax, var_6C
  loc_0066227F: lea ecx, var_14
  loc_00662282: push edx
  loc_00662283: push eax
  loc_00662284: mov var_C4, ecx
  loc_0066228A: mov var_CC, 00004008h
  loc_00662294: call [004011F8h] ; rtcHexVarFromVar
  loc_0066229A: lea ecx, var_6C
  loc_0066229D: push ecx
  loc_0066229E: call __vbaStrVarMove
  loc_006622A0: mov edx, eax
  loc_006622A2: lea ecx, var_14
  loc_006622A5: call ebx
  loc_006622A7: lea ecx, var_6C
  loc_006622AA: call [00401020h] ; __vbaFreeVar
  loc_006622B0: mov edx, var_14
  loc_006622B3: push edx
  loc_006622B4: call edi
  loc_006622B6: cmp eax, 00000008h
  loc_006622B9: jge 006622E0h
  loc_006622BB: mov esi, [00401070h] ; __vbaStrCat
  loc_006622C1: mov eax, var_14
  loc_006622C4: push eax
  loc_006622C5: call edi
  loc_006622C7: cmp eax, 00000008h
  loc_006622CA: jz 006622E0h
  loc_006622CC: mov ecx, var_14
  loc_006622CF: push 0042E3A4h
  loc_006622D4: push ecx
  loc_006622D5: call __vbaStrCat
  loc_006622D7: mov edx, eax
  loc_006622D9: lea ecx, var_14
  loc_006622DC: call ebx
  loc_006622DE: jmp 006622C1h
  loc_006622E0: mov edi, Me
  loc_006622E3: push edi
  loc_006622E4: mov edx, [edi]
  loc_006622E6: call [edx+00000704h]
  loc_006622EC: mov eax, var_14
  loc_006622EF: mov ecx, [edi+00000034h]
  loc_006622F2: push eax
  loc_006622F3: push ecx
  loc_006622F4: call [0040112Ch] ; __vbaStrCmp
  loc_006622FA: neg eax
  loc_006622FC: sbb eax, eax
  loc_006622FE: neg eax
  loc_00662300: dec eax
  loc_00662301: test ax, ax
  loc_00662304: jnz 006623EDh
  loc_0066230A: lea edx, var_6C
  loc_0066230D: push 0000000Dh
  loc_0066230F: push edx
  loc_00662310: call [004011B0h] ; rtcVarBstrFromAnsi
  loc_00662316: mov eax, 0000000Ah
  loc_0066231B: mov ecx, 80020004h
  loc_00662320: mov var_BC, eax
  loc_00662326: mov var_AC, eax
  loc_0066232C: mov var_9C, eax
  loc_00662332: mov eax, 00000008h
  loc_00662337: mov var_CC, eax
  loc_0066233D: mov var_DC, eax
  loc_00662343: mov var_B4, ecx
  loc_00662349: mov var_A4, ecx
  loc_0066234F: mov var_94, ecx
  loc_00662355: lea eax, var_BC
  loc_0066235B: mov esi, [004011C4h] ; __vbaVarCat
  loc_00662361: lea ecx, var_AC
  loc_00662367: push eax
  loc_00662368: lea edx, var_9C
  loc_0066236E: push ecx
  loc_0066236F: push edx
  loc_00662370: lea eax, var_CC
  loc_00662376: push 00000010h
  loc_00662378: lea ecx, var_6C
  loc_0066237B: push eax
  loc_0066237C: lea edx, var_7C
  loc_0066237F: push ecx
  loc_00662380: push edx
  loc_00662381: mov var_C4, 0043FE1Ch ; "You must provide a valid Machine ID in order to register this product."
  loc_0066238B: mov var_D4, 0043FEB0h ; "The application will now exit."
  loc_00662395: call __vbaVarCat
  loc_00662397: push eax
  loc_00662398: lea eax, var_DC
  loc_0066239E: lea ecx, var_8C
  loc_006623A4: push eax
  loc_006623A5: push ecx
  loc_006623A6: call __vbaVarCat
  loc_006623A8: push eax
  loc_006623A9: call [004010B4h] ; rtcMsgBox
  loc_006623AF: lea edx, var_BC
  loc_006623B5: lea eax, var_AC
  loc_006623BB: push edx
  loc_006623BC: lea ecx, var_9C
  loc_006623C2: push eax
  loc_006623C3: lea edx, var_8C
  loc_006623C9: push ecx
  loc_006623CA: lea eax, var_7C
  loc_006623CD: push edx
  loc_006623CE: lea ecx, var_6C
  loc_006623D1: push eax
  loc_006623D2: push ecx
  loc_006623D3: push 00000006h
  loc_006623D5: call [0040103Ch] ; __vbaFreeVarList
  loc_006623DB: mov edx, [edi]
  loc_006623DD: add esp, 0000001Ch
  loc_006623E0: push edi
  loc_006623E1: call [edx+00000718h]
  loc_006623E7: call [00401038h] ; __vbaEnd
  loc_006623ED: fwait
  loc_006623EE: push 0066246Fh
  loc_006623F3: jmp 00662437h
  loc_006623F5: lea ecx, var_58
  loc_006623F8: call [0040129Ch] ; __vbaFreeStr
  loc_006623FE: lea ecx, var_5C
  loc_00662401: call [004012A0h] ; __vbaFreeObj
  loc_00662407: lea eax, var_BC
  loc_0066240D: lea ecx, var_AC
  loc_00662413: push eax
  loc_00662414: lea edx, var_9C
  loc_0066241A: push ecx
  loc_0066241B: lea eax, var_8C
  loc_00662421: push edx
  loc_00662422: lea ecx, var_7C
  loc_00662425: push eax
  loc_00662426: lea edx, var_6C
  loc_00662429: push ecx
  loc_0066242A: push edx
  loc_0066242B: push 00000006h
  loc_0066242D: call [0040103Ch] ; __vbaFreeVarList
  loc_00662433: add esp, 0000001Ch
  loc_00662436: ret
  loc_00662437: lea ecx, var_14
  loc_0066243A: call [0040129Ch] ; __vbaFreeStr
  loc_00662440: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00662446: lea ecx, var_110
  loc_0066244C: lea eax, var_34
  loc_0066244F: push ecx
  loc_00662450: push 00000000h
  loc_00662452: mov var_110, eax
  loc_00662458: call __vbaAryDestruct
  loc_0066245A: lea eax, var_114
  loc_00662460: lea edx, var_50
  loc_00662463: push eax
  loc_00662464: push 00000000h
  loc_00662466: mov var_114, edx
  loc_0066246C: call __vbaAryDestruct
  loc_0066246E: ret
  loc_0066246F: mov ecx, var_10
  loc_00662472: pop edi
  loc_00662473: pop esi
  loc_00662474: xor eax, eax
  loc_00662476: mov fs:[00000000h], ecx
  loc_0066247D: pop ebx
  loc_0066247E: mov esp, ebp
  loc_00662480: pop ebp
  loc_00662481: retn 0004h
End Sub
