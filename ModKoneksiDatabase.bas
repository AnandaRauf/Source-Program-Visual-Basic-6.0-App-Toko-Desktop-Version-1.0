
Private Sub Proc_2_0_55B320
  loc_0055B320: push ebp
  loc_0055B321: mov ebp, esp
  loc_0055B323: sub esp, 00000014h
  loc_0055B326: push 00405696h ; __vbaExceptHandler
  loc_0055B32B: mov eax, fs:[00000000h]
  loc_0055B331: push eax
  loc_0055B332: mov fs:[00000000h], esp
  loc_0055B339: sub esp, 000000E8h
  loc_0055B33F: push ebx
  loc_0055B340: push esi
  loc_0055B341: push edi
  loc_0055B342: mov var_14, esp
  loc_0055B345: mov var_10, 00401368h
  loc_0055B34C: xor edi, edi
  loc_0055B34E: mov var_C, edi
  loc_0055B351: mov var_8, edi
  loc_0055B354: mov var_2C, edi
  loc_0055B357: mov var_34, edi
  loc_0055B35A: mov var_38, edi
  loc_0055B35D: mov var_3C, edi
  loc_0055B360: mov var_40, edi
  loc_0055B363: mov var_44, edi
  loc_0055B366: mov var_48, edi
  loc_0055B369: mov var_4C, edi
  loc_0055B36C: mov var_50, edi
  loc_0055B36F: mov var_60, edi
  loc_0055B372: mov var_70, edi
  loc_0055B375: mov var_80, edi
  loc_0055B378: mov var_90, edi
  loc_0055B37E: mov var_A0, edi
  loc_0055B384: mov var_B0, edi
  loc_0055B38A: mov var_D4, edi
  loc_0055B390: mov var_D8, edi
  loc_0055B396: mov var_F0, edi
  loc_0055B39C: call 0055B030h
  loc_0055B3A1: mov edx, 0042DEB0h ; "SELECT * FROM setting"
  loc_0055B3A6: mov ecx, 0066802Ch
  loc_0055B3AB: call [004011FCh] ; __vbaStrCopy
  loc_0055B3B1: cmp [00668028h], edi
  loc_0055B3B7: jnz 0055B3C9h
  loc_0055B3B9: push 00668028h
  loc_0055B3BE: push 0042DF30h
  loc_0055B3C3: call [004011E8h] ; __vbaNew2
  loc_0055B3C9: mov esi, [00668028h]
  loc_0055B3CF: cmp [00668024h], edi
  loc_0055B3D5: jnz 0055B3E7h
  loc_0055B3D7: push 00668024h
  loc_0055B3DC: push 0042DEFCh
  loc_0055B3E1: call [004011E8h] ; __vbaNew2
  loc_0055B3E7: mov eax, [00668024h]
  loc_0055B3EC: mov var_A8, eax
  loc_0055B3F2: mov ecx, 00000009h
  loc_0055B3F7: mov var_B0, ecx
  loc_0055B3FD: mov edx, [0066802Ch]
  loc_0055B403: mov var_98, edx
  loc_0055B409: mov var_A0, 00000008h
  loc_0055B413: mov edi, [esi]
  loc_0055B415: push FFFFFFFFh
  loc_0055B417: push 00000004h
  loc_0055B419: push 00000002h
  loc_0055B41B: sub esp, 00000010h
  loc_0055B41E: mov ebx, esp
  loc_0055B420: mov [ebx], ecx
  loc_0055B422: mov ecx, var_AC
  loc_0055B428: mov [ebx+00000004h], ecx
  loc_0055B42B: mov [ebx+00000008h], eax
  loc_0055B42E: mov eax, var_A4
  loc_0055B434: mov [ebx+0000000Ch], eax
  loc_0055B437: sub esp, 00000010h
  loc_0055B43A: mov ecx, esp
  loc_0055B43C: mov eax, var_A0
  loc_0055B442: mov [ecx], eax
  loc_0055B444: mov eax, var_9C
  loc_0055B44A: mov [ecx+00000004h], eax
  loc_0055B44D: mov [ecx+00000008h], edx
  loc_0055B450: mov edx, var_94
  loc_0055B456: mov [ecx+0000000Ch], edx
  loc_0055B459: push esi
  loc_0055B45A: call [edi+000000A0h]
  loc_0055B460: fnclex
  loc_0055B462: test eax, eax
  loc_0055B464: jge 0055B47Ch
  loc_0055B466: push 000000A0h
  loc_0055B46B: push 0042DF44h
  loc_0055B470: push esi
  loc_0055B471: push eax
  loc_0055B472: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0055B478: call edi
  loc_0055B47A: jmp 0055B482h
  loc_0055B47C: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0055B482: mov eax, [00668028h]
  loc_0055B487: test eax, eax
  loc_0055B489: jnz 0055B49Bh
  loc_0055B48B: push 00668028h
  loc_0055B490: push 0042DF30h
  loc_0055B495: call [004011E8h] ; __vbaNew2
  loc_0055B49B: mov esi, [00668028h]
  loc_0055B4A1: mov eax, [esi]
  loc_0055B4A3: lea ecx, var_D4
  loc_0055B4A9: push ecx
  loc_0055B4AA: push esi
  loc_0055B4AB: call [eax+00000034h]
  loc_0055B4AE: fnclex
  loc_0055B4B0: test eax, eax
  loc_0055B4B2: jge 0055B4BFh
  loc_0055B4B4: push 00000034h
  loc_0055B4B6: push 0042DF44h
  loc_0055B4BB: push esi
  loc_0055B4BC: push eax
  loc_0055B4BD: call edi
  loc_0055B4BF: cmp var_D4, 0000h
  loc_0055B4C7: jz 0055B56Eh
  loc_0055B4CD: mov edi, 80020004h
  loc_0055B4D2: mov var_88, edi
  loc_0055B4D8: mov ebx, 0000000Ah
  loc_0055B4DD: mov var_90, ebx
  loc_0055B4E3: mov var_78, edi
  loc_0055B4E6: mov var_80, ebx
  loc_0055B4E9: mov var_A8, 0042DF8Ch ; "Info"
  loc_0055B4F3: mov edi, 00000008h
  loc_0055B4F8: mov var_B0, edi
  loc_0055B4FE: lea edx, var_B0
  loc_0055B504: lea ecx, var_70
  loc_0055B507: mov esi, [00401238h] ; __vbaVarDup
  loc_0055B50D: call __vbaVarDup
  loc_0055B50F: mov var_98, 0042DF58h ; "Datbase setting kosong"
  loc_0055B519: mov var_A0, edi
  loc_0055B51F: lea edx, var_A0
  loc_0055B525: lea ecx, var_60
  loc_0055B528: call __vbaVarDup
  loc_0055B52A: lea edx, var_90
  loc_0055B530: push edx
  loc_0055B531: lea eax, var_80
  loc_0055B534: push eax
  loc_0055B535: lea ecx, var_70
  loc_0055B538: push ecx
  loc_0055B539: push 00000040h
  loc_0055B53B: lea edx, var_60
  loc_0055B53E: push edx
  loc_0055B53F: call [004010B4h] ; rtcMsgBox
  loc_0055B545: lea eax, var_90
  loc_0055B54B: push eax
  loc_0055B54C: lea ecx, var_80
  loc_0055B54F: push ecx
  loc_0055B550: lea edx, var_70
  loc_0055B553: push edx
  loc_0055B554: lea eax, var_60
  loc_0055B557: push eax
  loc_0055B558: push 00000004h
  loc_0055B55A: call [0040103Ch] ; __vbaFreeVarList
  loc_0055B560: add esp, 00000014h
  loc_0055B563: mov ebx, [0040126Ch] ; __vbaStrMove
  loc_0055B569: jmp 0055B728h
  loc_0055B56E: mov eax, [00668028h]
  loc_0055B573: test eax, eax
  loc_0055B575: jnz 0055B587h
  loc_0055B577: push 00668028h
  loc_0055B57C: push 0042DF30h
  loc_0055B581: call [004011E8h] ; __vbaNew2
  loc_0055B587: mov ecx, [00668028h]
  loc_0055B58D: push ecx
  loc_0055B58E: lea edx, var_F0
  loc_0055B594: push edx
  loc_0055B595: call [004010C8h] ; __vbaObjSetAddref
  loc_0055B59B: mov eax, var_F0
  loc_0055B5A1: mov ecx, [eax]
  loc_0055B5A3: lea edx, var_4C
  loc_0055B5A6: push edx
  loc_0055B5A7: push eax
  loc_0055B5A8: call [ecx+00000054h]
  loc_0055B5AB: fnclex
  loc_0055B5AD: test eax, eax
  loc_0055B5AF: jge 0055B5C2h
  loc_0055B5B1: push 00000054h
  loc_0055B5B3: push 0042DF44h
  loc_0055B5B8: mov ecx, var_F0
  loc_0055B5BE: push ecx
  loc_0055B5BF: push eax
  loc_0055B5C0: call edi
  loc_0055B5C2: mov eax, var_4C
  loc_0055B5C5: mov esi, eax
  loc_0055B5C7: mov ecx, 0042DF9Ch ; "Server"
  loc_0055B5CC: mov var_98, ecx
  loc_0055B5D2: mov edx, 00000008h
  loc_0055B5D7: mov var_A0, edx
  loc_0055B5DD: mov edi, [eax]
  loc_0055B5DF: lea ebx, var_50
  loc_0055B5E2: push ebx
  loc_0055B5E3: sub esp, 00000010h
  loc_0055B5E6: mov ebx, esp
  loc_0055B5E8: mov [ebx], edx
  loc_0055B5EA: mov edx, var_9C
  loc_0055B5F0: mov [ebx+00000004h], edx
  loc_0055B5F3: mov [ebx+00000008h], ecx
  loc_0055B5F6: mov ecx, var_94
  loc_0055B5FC: mov [ebx+0000000Ch], ecx
  loc_0055B5FF: push eax
  loc_0055B600: call [edi+00000028h]
  loc_0055B603: fnclex
  loc_0055B605: test eax, eax
  loc_0055B607: jge 0055B618h
  loc_0055B609: push 00000028h
  loc_0055B60B: push 0042DFACh
  loc_0055B610: push esi
  loc_0055B611: push eax
  loc_0055B612: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B618: mov eax, var_50
  loc_0055B61B: mov var_50, 00000000h
  loc_0055B622: mov var_58, eax
  loc_0055B625: mov var_60, 00000009h
  loc_0055B62C: lea edx, var_60
  loc_0055B62F: lea ecx, var_2C
  loc_0055B632: call [0040101Ch] ; __vbaVarMove
  loc_0055B638: lea ecx, var_4C
  loc_0055B63B: call [004012A0h] ; __vbaFreeObj
  loc_0055B641: mov eax, var_F0
  loc_0055B647: mov edx, [eax]
  loc_0055B649: lea ecx, var_4C
  loc_0055B64C: push ecx
  loc_0055B64D: push eax
  loc_0055B64E: call [edx+00000054h]
  loc_0055B651: fnclex
  loc_0055B653: test eax, eax
  loc_0055B655: jge 0055B66Ch
  loc_0055B657: push 00000054h
  loc_0055B659: push 0042DF44h
  loc_0055B65E: mov edx, var_F0
  loc_0055B664: push edx
  loc_0055B665: push eax
  loc_0055B666: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B66C: mov eax, var_4C
  loc_0055B66F: mov esi, eax
  loc_0055B671: mov ecx, 0042DFE0h ; "DBS"
  loc_0055B676: mov var_98, ecx
  loc_0055B67C: mov edx, 00000008h
  loc_0055B681: mov var_A0, edx
  loc_0055B687: mov edi, [eax]
  loc_0055B689: lea ebx, var_50
  loc_0055B68C: push ebx
  loc_0055B68D: sub esp, 00000010h
  loc_0055B690: mov ebx, esp
  loc_0055B692: mov [ebx], edx
  loc_0055B694: mov edx, var_9C
  loc_0055B69A: mov [ebx+00000004h], edx
  loc_0055B69D: mov [ebx+00000008h], ecx
  loc_0055B6A0: mov ecx, var_94
  loc_0055B6A6: mov [ebx+0000000Ch], ecx
  loc_0055B6A9: push eax
  loc_0055B6AA: call [edi+00000028h]
  loc_0055B6AD: fnclex
  loc_0055B6AF: test eax, eax
  loc_0055B6B1: jge 0055B6C2h
  loc_0055B6B3: push 00000028h
  loc_0055B6B5: push 0042DFACh
  loc_0055B6BA: push esi
  loc_0055B6BB: push eax
  loc_0055B6BC: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B6C2: mov eax, var_50
  loc_0055B6C5: mov esi, eax
  loc_0055B6C7: mov edx, [eax]
  loc_0055B6C9: lea ecx, var_60
  loc_0055B6CC: push ecx
  loc_0055B6CD: push eax
  loc_0055B6CE: call [edx+00000034h]
  loc_0055B6D1: fnclex
  loc_0055B6D3: test eax, eax
  loc_0055B6D5: jge 0055B6E6h
  loc_0055B6D7: push 00000034h
  loc_0055B6D9: push 0042DFBCh
  loc_0055B6DE: push esi
  loc_0055B6DF: push eax
  loc_0055B6E0: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B6E6: lea edx, var_60
  loc_0055B6E9: push edx
  loc_0055B6EA: call [00401034h] ; __vbaStrVarMove
  loc_0055B6F0: mov edx, eax
  loc_0055B6F2: lea ecx, var_34
  loc_0055B6F5: mov ebx, [0040126Ch] ; __vbaStrMove
  loc_0055B6FB: call ebx
  loc_0055B6FD: lea eax, var_50
  loc_0055B700: push eax
  loc_0055B701: lea ecx, var_4C
  loc_0055B704: push ecx
  loc_0055B705: push 00000002h
  loc_0055B707: call [0040104Ch] ; __vbaFreeObjList
  loc_0055B70D: add esp, 0000000Ch
  loc_0055B710: lea ecx, var_60
  loc_0055B713: call [00401020h] ; __vbaFreeVar
  loc_0055B719: push 00000000h
  loc_0055B71B: lea edx, var_F0
  loc_0055B721: push edx
  loc_0055B722: call [004010C8h] ; __vbaObjSetAddref
  loc_0055B728: mov edi, [004011E8h] ; __vbaNew2
  loc_0055B72E: push 00000001h
  loc_0055B730: call [004010BCh] ; __vbaOnError
  loc_0055B736: mov eax, [0066803Ch]
  loc_0055B73B: test eax, eax
  loc_0055B73D: jnz 0055B74Bh
  loc_0055B73F: push 0066803Ch
  loc_0055B744: push 0042DEFCh
  loc_0055B749: call edi
  loc_0055B74B: mov esi, [0066803Ch]
  loc_0055B751: mov eax, [esi]
  loc_0055B753: lea ecx, var_D8
  loc_0055B759: push ecx
  loc_0055B75A: push esi
  loc_0055B75B: call [eax+00000088h]
  loc_0055B761: fnclex
  loc_0055B763: test eax, eax
  loc_0055B765: jge 0055B779h
  loc_0055B767: push 00000088h
  loc_0055B76C: push 0042E1B0h
  loc_0055B771: push esi
  loc_0055B772: push eax
  loc_0055B773: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B779: cmp var_D8, 00000001h
  loc_0055B780: jnz 0055B7B8h
  loc_0055B782: mov eax, [0066803Ch]
  loc_0055B787: test eax, eax
  loc_0055B789: jnz 0055B797h
  loc_0055B78B: push 0066803Ch
  loc_0055B790: push 0042DEFCh
  loc_0055B795: call edi
  loc_0055B797: mov esi, [0066803Ch]
  loc_0055B79D: mov edx, [esi]
  loc_0055B79F: push esi
  loc_0055B7A0: call [edx+0000003Ch]
  loc_0055B7A3: fnclex
  loc_0055B7A5: test eax, eax
  loc_0055B7A7: jge 0055B7B8h
  loc_0055B7A9: push 0000003Ch
  loc_0055B7AB: push 0042E1B0h
  loc_0055B7B0: push esi
  loc_0055B7B1: push eax
  loc_0055B7B2: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B7B8: mov eax, [0066803Ch]
  loc_0055B7BD: test eax, eax
  loc_0055B7BF: jnz 0055B7CDh
  loc_0055B7C1: push 0066803Ch
  loc_0055B7C6: push 0042DEFCh
  loc_0055B7CB: call edi
  loc_0055B7CD: mov edi, [0066803Ch]
  loc_0055B7D3: push 0042E220h ; "provider=SQLOLEDB.1;"
  loc_0055B7D8: push 0042E250h ; "Integrated Security=SSPI;"
  loc_0055B7DD: mov esi, [00401070h] ; __vbaStrCat
  loc_0055B7E3: call __vbaStrCat
  loc_0055B7E5: mov edx, eax
  loc_0055B7E7: lea ecx, var_38
  loc_0055B7EA: call ebx
  loc_0055B7EC: push eax
  loc_0055B7ED: push 0042E288h ; "initial catalog="
  loc_0055B7F2: call __vbaStrCat
  loc_0055B7F4: mov edx, eax
  loc_0055B7F6: lea ecx, var_3C
  loc_0055B7F9: call ebx
  loc_0055B7FB: push eax
  loc_0055B7FC: mov eax, var_34
  loc_0055B7FF: push eax
  loc_0055B800: call __vbaStrCat
  loc_0055B802: mov edx, eax
  loc_0055B804: lea ecx, var_40
  loc_0055B807: call ebx
  loc_0055B809: push eax
  loc_0055B80A: push 0042E2B0h
  loc_0055B80F: call __vbaStrCat
  loc_0055B811: mov edx, eax
  loc_0055B813: lea ecx, var_44
  loc_0055B816: call ebx
  loc_0055B818: push eax
  loc_0055B819: push 0042E2B8h ; "Data Source="
  loc_0055B81E: call __vbaStrCat
  loc_0055B820: mov var_58, eax
  loc_0055B823: mov eax, 00000008h
  loc_0055B828: mov var_60, eax
  loc_0055B82B: mov var_98, 0042E2B0h
  loc_0055B835: mov var_A0, eax
  loc_0055B83B: mov ebx, [edi]
  loc_0055B83D: push FFFFFFFFh
  loc_0055B83F: push 0042DFECh
  loc_0055B844: push 0042DFECh
  loc_0055B849: lea ecx, var_60
  loc_0055B84C: push ecx
  loc_0055B84D: lea edx, var_2C
  loc_0055B850: push edx
  loc_0055B851: lea eax, var_70
  loc_0055B854: push eax
  loc_0055B855: mov esi, [004011C4h] ; __vbaVarCat
  loc_0055B85B: call __vbaVarCat
  loc_0055B85D: push eax
  loc_0055B85E: lea ecx, var_A0
  loc_0055B864: push ecx
  loc_0055B865: lea edx, var_80
  loc_0055B868: push edx
  loc_0055B869: call __vbaVarCat
  loc_0055B86B: push eax
  loc_0055B86C: lea eax, var_48
  loc_0055B86F: push eax
  loc_0055B870: call [004011C0h] ; __vbaStrVarVal
  loc_0055B876: push eax
  loc_0055B877: push edi
  loc_0055B878: call [ebx+00000050h]
  loc_0055B87B: fnclex
  loc_0055B87D: test eax, eax
  loc_0055B87F: jge 0055B890h
  loc_0055B881: push 00000050h
  loc_0055B883: push 0042E1B0h
  loc_0055B888: push edi
  loc_0055B889: push eax
  loc_0055B88A: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B890: lea ecx, var_48
  loc_0055B893: push ecx
  loc_0055B894: lea edx, var_44
  loc_0055B897: push edx
  loc_0055B898: lea eax, var_40
  loc_0055B89B: push eax
  loc_0055B89C: lea ecx, var_3C
  loc_0055B89F: push ecx
  loc_0055B8A0: lea edx, var_38
  loc_0055B8A3: push edx
  loc_0055B8A4: push 00000005h
  loc_0055B8A6: call [00401204h] ; __vbaFreeStrList
  loc_0055B8AC: lea eax, var_80
  loc_0055B8AF: push eax
  loc_0055B8B0: lea ecx, var_70
  loc_0055B8B3: push ecx
  loc_0055B8B4: lea edx, var_60
  loc_0055B8B7: push edx
  loc_0055B8B8: push 00000003h
  loc_0055B8BA: call [0040103Ch] ; __vbaFreeVarList
  loc_0055B8C0: add esp, 00000028h
  loc_0055B8C3: mov var_30, FFFFFFFFh
  loc_0055B8CA: call [004010A4h] ; __vbaExitProc
  loc_0055B8D0: push 0055BA8Dh
  loc_0055B8D5: jmp 0055BA6Eh
  loc_0055B8DA: mov var_30, 00000000h
  loc_0055B8E1: mov edi, 80020004h
  loc_0055B8E6: mov var_88, edi
  loc_0055B8EC: mov ebx, 0000000Ah
  loc_0055B8F1: mov var_90, ebx
  loc_0055B8F7: mov var_78, edi
  loc_0055B8FA: mov var_80, ebx
  loc_0055B8FD: mov var_A8, 0042E624h ; "Error"
  loc_0055B907: mov var_B0, 00000008h
  loc_0055B911: lea edx, var_B0
  loc_0055B917: lea ecx, var_70
  loc_0055B91A: mov esi, [00401238h] ; __vbaVarDup
  loc_0055B920: call __vbaVarDup
  loc_0055B922: mov var_98, 0042E650h ; "GAGAL KONEKSI SERVER, SILAHKAN PERBAIKI DATA SETTING ANDA !!"
  loc_0055B92C: mov var_A0, 00000008h
  loc_0055B936: lea edx, var_A0
  loc_0055B93C: lea ecx, var_60
  loc_0055B93F: call __vbaVarDup
  loc_0055B941: lea eax, var_90
  loc_0055B947: push eax
  loc_0055B948: lea ecx, var_80
  loc_0055B94B: push ecx
  loc_0055B94C: lea edx, var_70
  loc_0055B94F: push edx
  loc_0055B950: push 00000010h
  loc_0055B952: lea eax, var_60
  loc_0055B955: push eax
  loc_0055B956: call [004010B4h] ; rtcMsgBox
  loc_0055B95C: lea ecx, var_90
  loc_0055B962: push ecx
  loc_0055B963: lea edx, var_80
  loc_0055B966: push edx
  loc_0055B967: lea eax, var_70
  loc_0055B96A: push eax
  loc_0055B96B: lea ecx, var_60
  loc_0055B96E: push ecx
  loc_0055B96F: push 00000004h
  loc_0055B971: call [0040103Ch] ; __vbaFreeVarList
  loc_0055B977: add esp, 00000014h
  loc_0055B97A: mov eax, [00668010h]
  loc_0055B97F: test eax, eax
  loc_0055B981: jnz 0055B993h
  loc_0055B983: push 00668010h
  loc_0055B988: push 0040E2ECh
  loc_0055B98D: call [004011E8h] ; __vbaNew2
  loc_0055B993: mov esi, [00668010h]
  loc_0055B999: mov eax, edi
  loc_0055B99B: mov var_A8, eax
  loc_0055B9A1: mov ecx, ebx
  loc_0055B9A3: mov var_B0, ecx
  loc_0055B9A9: mov edx, edi
  loc_0055B9AB: mov var_98, edx
  loc_0055B9B1: mov var_A0, ebx
  loc_0055B9B7: mov edi, [esi]
  loc_0055B9B9: sub esp, 00000010h
  loc_0055B9BC: mov ebx, esp
  loc_0055B9BE: mov [ebx], ecx
  loc_0055B9C0: mov ecx, var_AC
  loc_0055B9C6: mov [ebx+00000004h], ecx
  loc_0055B9C9: mov [ebx+00000008h], eax
  loc_0055B9CC: mov eax, var_A4
  loc_0055B9D2: mov [ebx+0000000Ch], eax
  loc_0055B9D5: sub esp, 00000010h
  loc_0055B9D8: mov ecx, esp
  loc_0055B9DA: mov eax, var_A0
  loc_0055B9E0: mov [ecx], eax
  loc_0055B9E2: mov eax, var_9C
  loc_0055B9E8: mov [ecx+00000004h], eax
  loc_0055B9EB: mov [ecx+00000008h], edx
  loc_0055B9EE: mov edx, var_94
  loc_0055B9F4: mov [ecx+0000000Ch], edx
  loc_0055B9F7: push esi
  loc_0055B9F8: call [edi+000002B0h]
  loc_0055B9FE: fnclex
  loc_0055BA00: test eax, eax
  loc_0055BA02: jge 0055BA16h
  loc_0055BA04: push 000002B0h
  loc_0055BA09: push 0042DD20h
  loc_0055BA0E: push esi
  loc_0055BA0F: push eax
  loc_0055BA10: call [00401080h] ; __vbaHresultCheckObj
  loc_0055BA16: call [004010A4h] ; __vbaExitProc
  loc_0055BA1C: push 0055BA8Dh
  loc_0055BA21: jmp 0055BA6Eh
  loc_0055BA23: lea eax, var_48
  loc_0055BA26: push eax
  loc_0055BA27: lea ecx, var_44
  loc_0055BA2A: push ecx
  loc_0055BA2B: lea edx, var_40
  loc_0055BA2E: push edx
  loc_0055BA2F: lea eax, var_3C
  loc_0055BA32: push eax
  loc_0055BA33: lea ecx, var_38
  loc_0055BA36: push ecx
  loc_0055BA37: push 00000005h
  loc_0055BA39: call [00401204h] ; __vbaFreeStrList
  loc_0055BA3F: lea edx, var_50
  loc_0055BA42: push edx
  loc_0055BA43: lea eax, var_4C
  loc_0055BA46: push eax
  loc_0055BA47: push 00000002h
  loc_0055BA49: call [0040104Ch] ; __vbaFreeObjList
  loc_0055BA4F: lea ecx, var_90
  loc_0055BA55: push ecx
  loc_0055BA56: lea edx, var_80
  loc_0055BA59: push edx
  loc_0055BA5A: lea eax, var_70
  loc_0055BA5D: push eax
  loc_0055BA5E: lea ecx, var_60
  loc_0055BA61: push ecx
  loc_0055BA62: push 00000004h
  loc_0055BA64: call [0040103Ch] ; __vbaFreeVarList
  loc_0055BA6A: add esp, 00000038h
  loc_0055BA6D: ret
  loc_0055BA6E: lea ecx, var_F0
  loc_0055BA74: call [004012A0h] ; __vbaFreeObj
  loc_0055BA7A: lea ecx, var_2C
  loc_0055BA7D: call [00401020h] ; __vbaFreeVar
  loc_0055BA83: lea ecx, var_34
  loc_0055BA86: call [0040129Ch] ; __vbaFreeStr
  loc_0055BA8C: ret
  loc_0055BA8D: mov ax, var_30
  loc_0055BA91: mov ecx, var_1C
  loc_0055BA94: mov fs:[00000000h], ecx
  loc_0055BA9B: pop edi
  loc_0055BA9C: pop esi
  loc_0055BA9D: pop ebx
  loc_0055BA9E: mov esp, ebp
  loc_0055BAA0: pop ebp
  loc_0055BAA1: ret
  loc_0055BAA2: nop
End Sub
