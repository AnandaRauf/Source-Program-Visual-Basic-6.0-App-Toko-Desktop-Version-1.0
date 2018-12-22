VERSION 5.00
Begin VB.Form frmMenu
  Caption = "Form"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 225
  ClientTop = 855
  ClientWidth = 10470
  ClientHeight = 5325
  StartUpPosition = 3 'Windows Default
  Begin VB.Menu MuBeli
    Caption = "Data Pembelian"
    Begin VB.Menu SmGagal
      Caption = "Gagalkan Record Ini"
    End
  End
  Begin VB.Menu MuJual
    Caption = "Data Penjualan"
    Begin VB.Menu SmGagal2
      Caption = "Gagalkan Record Ini"
    End
  End
End

Attribute VB_Name = "frmMenu"


Private Sub SmGagal_Click() '5BB160
  loc_005BB160: push ebp
  loc_005BB161: mov ebp, esp
  loc_005BB163: sub esp, 00000018h
  loc_005BB166: push 00405696h ; __vbaExceptHandler
  loc_005BB16B: mov eax, fs:[00000000h]
  loc_005BB171: push eax
  loc_005BB172: mov fs:[00000000h], esp
  loc_005BB179: mov eax, 00000090h
  loc_005BB17E: call 00405690h ; __vbaChkstk
  loc_005BB183: push ebx
  loc_005BB184: push esi
  loc_005BB185: push edi
  loc_005BB186: mov var_18, esp
  loc_005BB189: mov var_14, 00403450h ; "'"
  loc_005BB190: mov eax, Me
  loc_005BB193: and eax, 00000001h
  loc_005BB196: mov var_10, eax
  loc_005BB199: mov ecx, Me
  loc_005BB19C: and ecx, FFFFFFFEh
  loc_005BB19F: mov Me, ecx
  loc_005BB1A2: mov var_C, 00000000h
  loc_005BB1A9: mov edx, Me
  loc_005BB1AC: mov eax, [edx]
  loc_005BB1AE: mov ecx, Me
  loc_005BB1B1: push ecx
  loc_005BB1B2: call [eax+00000004h]
  loc_005BB1B5: mov var_4, 00000001h
  loc_005BB1BC: mov var_4, 00000002h
  loc_005BB1C3: cmp [0066823Ch], 00000000h
  loc_005BB1CA: jnz 005BB1E8h
  loc_005BB1CC: push 0066823Ch
  loc_005BB1D1: push 00427B78h
  loc_005BB1D6: call [004011E8h] ; __vbaNew2
  loc_005BB1DC: mov var_A4, 0066823Ch
  loc_005BB1E6: jmp 005BB1F2h
  loc_005BB1E8: mov var_A4, 0066823Ch
  loc_005BB1F2: mov edx, var_A4
  loc_005BB1F8: mov eax, [edx]
  loc_005BB1FA: push eax
  loc_005BB1FB: lea ecx, var_8C
  loc_005BB201: push ecx
  loc_005BB202: call [004010C8h] ; __vbaObjSetAddref
  loc_005BB208: mov var_4, 00000003h
  loc_005BB20F: push 00000000h
  loc_005BB211: push 0000000Ah
  loc_005BB213: mov edx, var_8C
  loc_005BB219: mov eax, [edx]
  loc_005BB21B: mov ecx, var_8C
  loc_005BB221: push ecx
  loc_005BB222: call [eax+00000390h]
  loc_005BB228: push eax
  loc_005BB229: lea edx, var_24
  loc_005BB22C: push edx
  loc_005BB22D: call [004010B8h] ; __vbaObjSet
  loc_005BB233: push eax
  loc_005BB234: lea eax, var_34
  loc_005BB237: push eax
  loc_005BB238: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB23E: add esp, 00000010h
  loc_005BB241: push eax
  loc_005BB242: call [0040121Ch] ; __vbaI4Var
  loc_005BB248: mov ecx, eax
  loc_005BB24A: call [00401138h] ; __vbaI2I4
  loc_005BB250: mov ecx, Me
  loc_005BB253: mov [ecx+00000034h], ax
  loc_005BB257: lea ecx, var_24
  loc_005BB25A: call [004012A0h] ; __vbaFreeObj
  loc_005BB260: lea ecx, var_34
  loc_005BB263: call [00401020h] ; __vbaFreeVar
  loc_005BB269: mov var_4, 00000004h
  loc_005BB270: mov edx, Me
  loc_005BB273: movsx eax, [edx+00000034h]
  loc_005BB277: mov var_4C, eax
  loc_005BB27A: mov var_54, 00000003h
  loc_005BB281: mov var_6C, 00000000h
  loc_005BB288: mov var_74, 00000003h
  loc_005BB28F: mov eax, 00000010h
  loc_005BB294: call 00405690h ; __vbaChkstk
  loc_005BB299: mov ecx, esp
  loc_005BB29B: mov edx, var_54
  loc_005BB29E: mov [ecx], edx
  loc_005BB2A0: mov eax, var_50
  loc_005BB2A3: mov [ecx+00000004h], eax
  loc_005BB2A6: mov edx, var_4C
  loc_005BB2A9: mov [ecx+00000008h], edx
  loc_005BB2AC: mov eax, var_48
  loc_005BB2AF: mov [ecx+0000000Ch], eax
  loc_005BB2B2: mov eax, 00000010h
  loc_005BB2B7: call 00405690h ; __vbaChkstk
  loc_005BB2BC: mov ecx, esp
  loc_005BB2BE: mov edx, var_74
  loc_005BB2C1: mov [ecx], edx
  loc_005BB2C3: mov eax, var_70
  loc_005BB2C6: mov [ecx+00000004h], eax
  loc_005BB2C9: mov edx, var_6C
  loc_005BB2CC: mov [ecx+00000008h], edx
  loc_005BB2CF: mov eax, var_68
  loc_005BB2D2: mov [ecx+0000000Ch], eax
  loc_005BB2D5: push 00000002h
  loc_005BB2D7: push 00000041h
  loc_005BB2D9: mov ecx, var_8C
  loc_005BB2DF: mov edx, [ecx]
  loc_005BB2E1: mov eax, var_8C
  loc_005BB2E7: push eax
  loc_005BB2E8: call [edx+00000390h]
  loc_005BB2EE: push eax
  loc_005BB2EF: lea ecx, var_24
  loc_005BB2F2: push ecx
  loc_005BB2F3: call [004010B8h] ; __vbaObjSet
  loc_005BB2F9: push eax
  loc_005BB2FA: lea edx, var_34
  loc_005BB2FD: push edx
  loc_005BB2FE: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB304: add esp, 00000030h
  loc_005BB307: push eax
  loc_005BB308: call [00401034h] ; __vbaStrVarMove
  loc_005BB30E: mov var_3C, eax
  loc_005BB311: mov var_44, 00000008h
  loc_005BB318: lea edx, var_44
  loc_005BB31B: mov ecx, Me
  loc_005BB31E: add ecx, 00000038h
  loc_005BB321: call [0040101Ch] ; __vbaVarMove
  loc_005BB327: lea ecx, var_24
  loc_005BB32A: call [004012A0h] ; __vbaFreeObj
  loc_005BB330: lea eax, var_44
  loc_005BB333: push eax
  loc_005BB334: lea ecx, var_34
  loc_005BB337: push ecx
  loc_005BB338: push 00000002h
  loc_005BB33A: call [0040103Ch] ; __vbaFreeVarList
  loc_005BB340: add esp, 0000000Ch
  loc_005BB343: mov var_4, 00000005h
  loc_005BB34A: push 00000000h
  loc_005BB34C: push 00000004h
  loc_005BB34E: mov edx, var_8C
  loc_005BB354: mov eax, [edx]
  loc_005BB356: mov ecx, var_8C
  loc_005BB35C: push ecx
  loc_005BB35D: call [eax+00000390h]
  loc_005BB363: push eax
  loc_005BB364: lea edx, var_24
  loc_005BB367: push edx
  loc_005BB368: call [004010B8h] ; __vbaObjSet
  loc_005BB36E: push eax
  loc_005BB36F: lea eax, var_34
  loc_005BB372: push eax
  loc_005BB373: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB379: add esp, 00000010h
  loc_005BB37C: push eax
  loc_005BB37D: call [0040121Ch] ; __vbaI4Var
  loc_005BB383: xor ecx, ecx
  loc_005BB385: cmp eax, 00000002h
  loc_005BB388: setl cl
  loc_005BB38B: neg ecx
  loc_005BB38D: mov var_88, cx
  loc_005BB394: lea ecx, var_24
  loc_005BB397: call [004012A0h] ; __vbaFreeObj
  loc_005BB39D: lea ecx, var_34
  loc_005BB3A0: call [00401020h] ; __vbaFreeVar
  loc_005BB3A6: movsx edx, var_88
  loc_005BB3AD: test edx, edx
  loc_005BB3AF: jz 005BB474h
  loc_005BB3B5: mov var_4, 00000006h
  loc_005BB3BC: mov eax, var_8C
  loc_005BB3C2: mov ecx, [eax]
  loc_005BB3C4: mov edx, var_8C
  loc_005BB3CA: push edx
  loc_005BB3CB: call [ecx+000006FCh]
  loc_005BB3D1: fnclex
  loc_005BB3D3: mov var_88, eax
  loc_005BB3D9: cmp var_88, 00000000h
  loc_005BB3E0: jge 005BB408h
  loc_005BB3E2: push 000006FCh
  loc_005BB3E7: push 00433DD8h
  loc_005BB3EC: mov eax, var_8C
  loc_005BB3F2: push eax
  loc_005BB3F3: mov ecx, var_88
  loc_005BB3F9: push ecx
  loc_005BB3FA: call [00401080h] ; __vbaHresultCheckObj
  loc_005BB400: mov var_A8, eax
  loc_005BB406: jmp 005BB412h
  loc_005BB408: mov var_A8, 00000000h
  loc_005BB412: mov var_4, 00000007h
  loc_005BB419: mov edx, var_8C
  loc_005BB41F: mov eax, [edx]
  loc_005BB421: mov ecx, var_8C
  loc_005BB427: push ecx
  loc_005BB428: call [eax+00000704h]
  loc_005BB42E: fnclex
  loc_005BB430: mov var_88, eax
  loc_005BB436: cmp var_88, 00000000h
  loc_005BB43D: jge 005BB465h
  loc_005BB43F: push 00000704h
  loc_005BB444: push 00433DD8h
  loc_005BB449: mov edx, var_8C
  loc_005BB44F: push edx
  loc_005BB450: mov eax, var_88
  loc_005BB456: push eax
  loc_005BB457: call [00401080h] ; __vbaHresultCheckObj
  loc_005BB45D: mov var_AC, eax
  loc_005BB463: jmp 005BB46Fh
  loc_005BB465: mov var_AC, 00000000h
  loc_005BB46F: jmp 005BB552h
  loc_005BB474: mov var_4, 00000009h
  loc_005BB47B: push FFFFFFFFh
  loc_005BB47D: call [004010BCh] ; __vbaOnError
  loc_005BB483: mov var_4, 0000000Ah
  loc_005BB48A: mov ecx, Me
  loc_005BB48D: movsx edx, [ecx+00000034h]
  loc_005BB491: mov var_4C, edx
  loc_005BB494: mov var_54, 00000003h
  loc_005BB49B: mov eax, 00000010h
  loc_005BB4A0: call 00405690h ; __vbaChkstk
  loc_005BB4A5: mov eax, esp
  loc_005BB4A7: mov ecx, var_54
  loc_005BB4AA: mov [eax], ecx
  loc_005BB4AC: mov edx, var_50
  loc_005BB4AF: mov [eax+00000004h], edx
  loc_005BB4B2: mov ecx, var_4C
  loc_005BB4B5: mov [eax+00000008h], ecx
  loc_005BB4B8: mov edx, var_48
  loc_005BB4BB: mov [eax+0000000Ch], edx
  loc_005BB4BE: push 00000001h
  loc_005BB4C0: push 00000043h
  loc_005BB4C2: mov eax, var_8C
  loc_005BB4C8: mov ecx, [eax]
  loc_005BB4CA: mov edx, var_8C
  loc_005BB4D0: push edx
  loc_005BB4D1: call [ecx+00000390h]
  loc_005BB4D7: push eax
  loc_005BB4D8: lea eax, var_24
  loc_005BB4DB: push eax
  loc_005BB4DC: call [004010B8h] ; __vbaObjSet
  loc_005BB4E2: push eax
  loc_005BB4E3: call [0040102Ch] ; __vbaLateIdCall
  loc_005BB4E9: add esp, 0000001Ch
  loc_005BB4EC: lea ecx, var_24
  loc_005BB4EF: call [004012A0h] ; __vbaFreeObj
  loc_005BB4F5: mov var_4, 0000000Bh
  loc_005BB4FC: mov ecx, var_8C
  loc_005BB502: mov edx, [ecx]
  loc_005BB504: mov eax, var_8C
  loc_005BB50A: push eax
  loc_005BB50B: call [edx+00000718h]
  loc_005BB511: fnclex
  loc_005BB513: mov var_88, eax
  loc_005BB519: cmp var_88, 00000000h
  loc_005BB520: jge 005BB548h
  loc_005BB522: push 00000718h
  loc_005BB527: push 00433DD8h
  loc_005BB52C: mov ecx, var_8C
  loc_005BB532: push ecx
  loc_005BB533: mov edx, var_88
  loc_005BB539: push edx
  loc_005BB53A: call [00401080h] ; __vbaHresultCheckObj
  loc_005BB540: mov var_B0, eax
  loc_005BB546: jmp 005BB552h
  loc_005BB548: mov var_B0, 00000000h
  loc_005BB552: mov var_4, 0000000Dh
  loc_005BB559: push 00000000h
  loc_005BB55B: push FFFFFDDAh
  loc_005BB560: mov eax, var_8C
  loc_005BB566: mov ecx, [eax]
  loc_005BB568: mov edx, var_8C
  loc_005BB56E: push edx
  loc_005BB56F: call [ecx+00000390h]
  loc_005BB575: push eax
  loc_005BB576: lea eax, var_24
  loc_005BB579: push eax
  loc_005BB57A: call [004010B8h] ; __vbaObjSet
  loc_005BB580: push eax
  loc_005BB581: call [0040102Ch] ; __vbaLateIdCall
  loc_005BB587: add esp, 0000000Ch
  loc_005BB58A: lea ecx, var_24
  loc_005BB58D: call [004012A0h] ; __vbaFreeObj
  loc_005BB593: mov var_4, 0000000Eh
  loc_005BB59A: push 00000000h
  loc_005BB59C: lea ecx, var_8C
  loc_005BB5A2: push ecx
  loc_005BB5A3: call [004010C8h] ; __vbaObjSetAddref
  loc_005BB5A9: mov var_10, 00000000h
  loc_005BB5B0: push 005BB5E1h
  loc_005BB5B5: jmp 005BB5D4h
  loc_005BB5B7: lea ecx, var_24
  loc_005BB5BA: call [004012A0h] ; __vbaFreeObj
  loc_005BB5C0: lea edx, var_44
  loc_005BB5C3: push edx
  loc_005BB5C4: lea eax, var_34
  loc_005BB5C7: push eax
  loc_005BB5C8: push 00000002h
  loc_005BB5CA: call [0040103Ch] ; __vbaFreeVarList
  loc_005BB5D0: add esp, 0000000Ch
  loc_005BB5D3: ret
  loc_005BB5D4: lea ecx, var_8C
  loc_005BB5DA: call [004012A0h] ; __vbaFreeObj
  loc_005BB5E0: ret
  loc_005BB5E1: mov ecx, Me
  loc_005BB5E4: mov edx, [ecx]
  loc_005BB5E6: mov eax, Me
  loc_005BB5E9: push eax
  loc_005BB5EA: call [edx+00000008h]
  loc_005BB5ED: mov eax, var_10
  loc_005BB5F0: mov ecx, var_20
  loc_005BB5F3: mov fs:[00000000h], ecx
  loc_005BB5FA: pop edi
  loc_005BB5FB: pop esi
  loc_005BB5FC: pop ebx
  loc_005BB5FD: mov esp, ebp
  loc_005BB5FF: pop ebp
  loc_005BB600: retn 0004h
End Sub

Private Sub SmGagal2_Click() '5BB610
  loc_005BB610: push ebp
  loc_005BB611: mov ebp, esp
  loc_005BB613: sub esp, 00000018h
  loc_005BB616: push 00405696h ; __vbaExceptHandler
  loc_005BB61B: mov eax, fs:[00000000h]
  loc_005BB621: push eax
  loc_005BB622: mov fs:[00000000h], esp
  loc_005BB629: mov eax, 00000088h
  loc_005BB62E: call 00405690h ; __vbaChkstk
  loc_005BB633: push ebx
  loc_005BB634: push esi
  loc_005BB635: push edi
  loc_005BB636: mov var_18, esp
  loc_005BB639: mov var_14, 004034B0h ; "'"
  loc_005BB640: mov eax, Me
  loc_005BB643: and eax, 00000001h
  loc_005BB646: mov var_10, eax
  loc_005BB649: mov ecx, Me
  loc_005BB64C: and ecx, FFFFFFFEh
  loc_005BB64F: mov Me, ecx
  loc_005BB652: mov var_C, 00000000h
  loc_005BB659: mov edx, Me
  loc_005BB65C: mov eax, [edx]
  loc_005BB65E: mov ecx, Me
  loc_005BB661: push ecx
  loc_005BB662: call [eax+00000004h]
  loc_005BB665: mov var_4, 00000001h
  loc_005BB66C: mov var_4, 00000002h
  loc_005BB673: cmp [00668278h], 00000000h
  loc_005BB67A: jnz 005BB698h
  loc_005BB67C: push 00668278h
  loc_005BB681: push 00424728h
  loc_005BB686: call [004011E8h] ; __vbaNew2
  loc_005BB68C: mov var_A4, 00668278h
  loc_005BB696: jmp 005BB6A2h
  loc_005BB698: mov var_A4, 00668278h
  loc_005BB6A2: mov edx, var_A4
  loc_005BB6A8: mov eax, [edx]
  loc_005BB6AA: push eax
  loc_005BB6AB: lea ecx, var_8C
  loc_005BB6B1: push ecx
  loc_005BB6B2: call [004010C8h] ; __vbaObjSetAddref
  loc_005BB6B8: mov var_4, 00000003h
  loc_005BB6BF: push 00000000h
  loc_005BB6C1: push 0000000Ah
  loc_005BB6C3: mov edx, var_8C
  loc_005BB6C9: mov eax, [edx]
  loc_005BB6CB: mov ecx, var_8C
  loc_005BB6D1: push ecx
  loc_005BB6D2: call [eax+00000390h]
  loc_005BB6D8: push eax
  loc_005BB6D9: lea edx, var_24
  loc_005BB6DC: push edx
  loc_005BB6DD: call [004010B8h] ; __vbaObjSet
  loc_005BB6E3: push eax
  loc_005BB6E4: lea eax, var_34
  loc_005BB6E7: push eax
  loc_005BB6E8: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB6EE: add esp, 00000010h
  loc_005BB6F1: push eax
  loc_005BB6F2: call [0040121Ch] ; __vbaI4Var
  loc_005BB6F8: mov ecx, eax
  loc_005BB6FA: call [00401138h] ; __vbaI2I4
  loc_005BB700: mov ecx, Me
  loc_005BB703: mov [ecx+00000034h], ax
  loc_005BB707: lea ecx, var_24
  loc_005BB70A: call [004012A0h] ; __vbaFreeObj
  loc_005BB710: lea ecx, var_34
  loc_005BB713: call [00401020h] ; __vbaFreeVar
  loc_005BB719: mov var_4, 00000004h
  loc_005BB720: mov edx, Me
  loc_005BB723: movsx eax, [edx+00000034h]
  loc_005BB727: mov var_4C, eax
  loc_005BB72A: mov var_54, 00000003h
  loc_005BB731: mov var_6C, 00000000h
  loc_005BB738: mov var_74, 00000003h
  loc_005BB73F: mov eax, 00000010h
  loc_005BB744: call 00405690h ; __vbaChkstk
  loc_005BB749: mov ecx, esp
  loc_005BB74B: mov edx, var_54
  loc_005BB74E: mov [ecx], edx
  loc_005BB750: mov eax, var_50
  loc_005BB753: mov [ecx+00000004h], eax
  loc_005BB756: mov edx, var_4C
  loc_005BB759: mov [ecx+00000008h], edx
  loc_005BB75C: mov eax, var_48
  loc_005BB75F: mov [ecx+0000000Ch], eax
  loc_005BB762: mov eax, 00000010h
  loc_005BB767: call 00405690h ; __vbaChkstk
  loc_005BB76C: mov ecx, esp
  loc_005BB76E: mov edx, var_74
  loc_005BB771: mov [ecx], edx
  loc_005BB773: mov eax, var_70
  loc_005BB776: mov [ecx+00000004h], eax
  loc_005BB779: mov edx, var_6C
  loc_005BB77C: mov [ecx+00000008h], edx
  loc_005BB77F: mov eax, var_68
  loc_005BB782: mov [ecx+0000000Ch], eax
  loc_005BB785: push 00000002h
  loc_005BB787: push 00000041h
  loc_005BB789: mov ecx, var_8C
  loc_005BB78F: mov edx, [ecx]
  loc_005BB791: mov eax, var_8C
  loc_005BB797: push eax
  loc_005BB798: call [edx+00000390h]
  loc_005BB79E: push eax
  loc_005BB79F: lea ecx, var_24
  loc_005BB7A2: push ecx
  loc_005BB7A3: call [004010B8h] ; __vbaObjSet
  loc_005BB7A9: push eax
  loc_005BB7AA: lea edx, var_34
  loc_005BB7AD: push edx
  loc_005BB7AE: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB7B4: add esp, 00000030h
  loc_005BB7B7: push eax
  loc_005BB7B8: call [00401034h] ; __vbaStrVarMove
  loc_005BB7BE: mov var_3C, eax
  loc_005BB7C1: mov var_44, 00000008h
  loc_005BB7C8: lea edx, var_44
  loc_005BB7CB: mov ecx, Me
  loc_005BB7CE: add ecx, 00000038h
  loc_005BB7D1: call [0040101Ch] ; __vbaVarMove
  loc_005BB7D7: lea ecx, var_24
  loc_005BB7DA: call [004012A0h] ; __vbaFreeObj
  loc_005BB7E0: lea eax, var_44
  loc_005BB7E3: push eax
  loc_005BB7E4: lea ecx, var_34
  loc_005BB7E7: push ecx
  loc_005BB7E8: push 00000002h
  loc_005BB7EA: call [0040103Ch] ; __vbaFreeVarList
  loc_005BB7F0: add esp, 0000000Ch
  loc_005BB7F3: mov var_4, 00000005h
  loc_005BB7FA: push 00000000h
  loc_005BB7FC: push 00000004h
  loc_005BB7FE: mov edx, var_8C
  loc_005BB804: mov eax, [edx]
  loc_005BB806: mov ecx, var_8C
  loc_005BB80C: push ecx
  loc_005BB80D: call [eax+00000390h]
  loc_005BB813: push eax
  loc_005BB814: lea edx, var_24
  loc_005BB817: push edx
  loc_005BB818: call [004010B8h] ; __vbaObjSet
  loc_005BB81E: push eax
  loc_005BB81F: lea eax, var_34
  loc_005BB822: push eax
  loc_005BB823: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB829: add esp, 00000010h
  loc_005BB82C: push eax
  loc_005BB82D: call [0040121Ch] ; __vbaI4Var
  loc_005BB833: xor ecx, ecx
  loc_005BB835: cmp eax, 00000002h
  loc_005BB838: setl cl
  loc_005BB83B: neg ecx
  loc_005BB83D: mov var_88, cx
  loc_005BB844: lea ecx, var_24
  loc_005BB847: call [004012A0h] ; __vbaFreeObj
  loc_005BB84D: lea ecx, var_34
  loc_005BB850: call [00401020h] ; __vbaFreeVar
  loc_005BB856: movsx edx, var_88
  loc_005BB85D: test edx, edx
  loc_005BB85F: jz 005BB866h
  loc_005BB861: jmp 005BB9C2h
  loc_005BB866: mov var_4, 00000007h
  loc_005BB86D: push FFFFFFFFh
  loc_005BB86F: call [004010BCh] ; __vbaOnError
  loc_005BB875: mov var_4, 00000008h
  loc_005BB87C: push 00000000h
  loc_005BB87E: push 00000004h
  loc_005BB880: mov eax, var_8C
  loc_005BB886: mov ecx, [eax]
  loc_005BB888: mov edx, var_8C
  loc_005BB88E: push edx
  loc_005BB88F: call [ecx+00000390h]
  loc_005BB895: push eax
  loc_005BB896: lea eax, var_24
  loc_005BB899: push eax
  loc_005BB89A: call [004010B8h] ; __vbaObjSet
  loc_005BB8A0: push eax
  loc_005BB8A1: lea ecx, var_34
  loc_005BB8A4: push ecx
  loc_005BB8A5: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005BB8AB: add esp, 00000010h
  loc_005BB8AE: push eax
  loc_005BB8AF: call [0040121Ch] ; __vbaI4Var
  loc_005BB8B5: xor edx, edx
  loc_005BB8B7: cmp eax, 00000002h
  loc_005BB8BA: setz dl
  loc_005BB8BD: neg edx
  loc_005BB8BF: mov var_88, dx
  loc_005BB8C6: lea ecx, var_24
  loc_005BB8C9: call [004012A0h] ; __vbaFreeObj
  loc_005BB8CF: lea ecx, var_34
  loc_005BB8D2: call [00401020h] ; __vbaFreeVar
  loc_005BB8D8: movsx eax, var_88
  loc_005BB8DF: test eax, eax
  loc_005BB8E1: jz 005BB8F3h
  loc_005BB8E3: mov var_4, 00000009h
  loc_005BB8EA: mov ecx, Me
  loc_005BB8ED: mov [ecx+00000034h], 0001h
  loc_005BB8F3: mov var_4, 0000000Ch
  loc_005BB8FA: mov edx, Me
  loc_005BB8FD: movsx eax, [edx+00000034h]
  loc_005BB901: mov var_4C, eax
  loc_005BB904: mov var_54, 00000003h
  loc_005BB90B: mov eax, 00000010h
  loc_005BB910: call 00405690h ; __vbaChkstk
  loc_005BB915: mov ecx, esp
  loc_005BB917: mov edx, var_54
  loc_005BB91A: mov [ecx], edx
  loc_005BB91C: mov eax, var_50
  loc_005BB91F: mov [ecx+00000004h], eax
  loc_005BB922: mov edx, var_4C
  loc_005BB925: mov [ecx+00000008h], edx
  loc_005BB928: mov eax, var_48
  loc_005BB92B: mov [ecx+0000000Ch], eax
  loc_005BB92E: push 00000001h
  loc_005BB930: push 00000043h
  loc_005BB932: mov ecx, var_8C
  loc_005BB938: mov edx, [ecx]
  loc_005BB93A: mov eax, var_8C
  loc_005BB940: push eax
  loc_005BB941: call [edx+00000390h]
  loc_005BB947: push eax
  loc_005BB948: lea ecx, var_24
  loc_005BB94B: push ecx
  loc_005BB94C: call [004010B8h] ; __vbaObjSet
  loc_005BB952: push eax
  loc_005BB953: call [0040102Ch] ; __vbaLateIdCall
  loc_005BB959: add esp, 0000001Ch
  loc_005BB95C: lea ecx, var_24
  loc_005BB95F: call [004012A0h] ; __vbaFreeObj
  loc_005BB965: mov var_4, 0000000Dh
  loc_005BB96C: mov edx, var_8C
  loc_005BB972: mov eax, [edx]
  loc_005BB974: mov ecx, var_8C
  loc_005BB97A: push ecx
  loc_005BB97B: call [eax+0000071Ch]
  loc_005BB981: fnclex
  loc_005BB983: mov var_88, eax
  loc_005BB989: cmp var_88, 00000000h
  loc_005BB990: jge 005BB9B8h
  loc_005BB992: push 0000071Ch
  loc_005BB997: push 00434B78h
  loc_005BB99C: mov edx, var_8C
  loc_005BB9A2: push edx
  loc_005BB9A3: mov eax, var_88
  loc_005BB9A9: push eax
  loc_005BB9AA: call [00401080h] ; __vbaHresultCheckObj
  loc_005BB9B0: mov var_A8, eax
  loc_005BB9B6: jmp 005BB9C2h
  loc_005BB9B8: mov var_A8, 00000000h
  loc_005BB9C2: mov var_4, 0000000Fh
  loc_005BB9C9: push 00000000h
  loc_005BB9CB: push FFFFFDDAh
  loc_005BB9D0: mov ecx, var_8C
  loc_005BB9D6: mov edx, [ecx]
  loc_005BB9D8: mov eax, var_8C
  loc_005BB9DE: push eax
  loc_005BB9DF: call [edx+00000390h]
  loc_005BB9E5: push eax
  loc_005BB9E6: lea ecx, var_24
  loc_005BB9E9: push ecx
  loc_005BB9EA: call [004010B8h] ; __vbaObjSet
  loc_005BB9F0: push eax
  loc_005BB9F1: call [0040102Ch] ; __vbaLateIdCall
  loc_005BB9F7: add esp, 0000000Ch
  loc_005BB9FA: lea ecx, var_24
  loc_005BB9FD: call [004012A0h] ; __vbaFreeObj
  loc_005BBA03: mov var_4, 00000010h
  loc_005BBA0A: push 00000000h
  loc_005BBA0C: lea edx, var_8C
  loc_005BBA12: push edx
  loc_005BBA13: call [004010C8h] ; __vbaObjSetAddref
  loc_005BBA19: mov var_10, 00000000h
  loc_005BBA20: push 005BBA51h
  loc_005BBA25: jmp 005BBA44h
  loc_005BBA27: lea ecx, var_24
  loc_005BBA2A: call [004012A0h] ; __vbaFreeObj
  loc_005BBA30: lea eax, var_44
  loc_005BBA33: push eax
  loc_005BBA34: lea ecx, var_34
  loc_005BBA37: push ecx
  loc_005BBA38: push 00000002h
  loc_005BBA3A: call [0040103Ch] ; __vbaFreeVarList
  loc_005BBA40: add esp, 0000000Ch
  loc_005BBA43: ret
  loc_005BBA44: lea ecx, var_8C
  loc_005BBA4A: call [004012A0h] ; __vbaFreeObj
  loc_005BBA50: ret
  loc_005BBA51: mov edx, Me
  loc_005BBA54: mov eax, [edx]
  loc_005BBA56: mov ecx, Me
  loc_005BBA59: push ecx
  loc_005BBA5A: call [eax+00000008h]
  loc_005BBA5D: mov eax, var_10
  loc_005BBA60: mov ecx, var_20
  loc_005BBA63: mov fs:[00000000h], ecx
  loc_005BBA6A: pop edi
  loc_005BBA6B: pop esi
  loc_005BBA6C: pop ebx
  loc_005BBA6D: mov esp, ebp
  loc_005BBA6F: pop ebp
  loc_005BBA70: retn 0004h
End Sub
