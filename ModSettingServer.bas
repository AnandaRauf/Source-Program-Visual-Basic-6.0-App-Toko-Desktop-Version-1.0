
Private Sub Proc_1_0_55B030
  loc_0055B030: push ebp
  loc_0055B031: mov ebp, esp
  loc_0055B033: sub esp, 00000014h
  loc_0055B036: push 00405696h ; __vbaExceptHandler
  loc_0055B03B: mov eax, fs:[00000000h]
  loc_0055B041: push eax
  loc_0055B042: mov fs:[00000000h], esp
  loc_0055B049: sub esp, 000000C0h
  loc_0055B04F: push ebx
  loc_0055B050: push esi
  loc_0055B051: push edi
  loc_0055B052: mov var_14, esp
  loc_0055B055: mov var_10, 00401340h
  loc_0055B05C: xor edi, edi
  loc_0055B05E: mov var_C, edi
  loc_0055B061: mov var_8, edi
  loc_0055B064: mov var_24, edi
  loc_0055B067: mov var_28, edi
  loc_0055B06A: mov var_2C, edi
  loc_0055B06D: mov var_30, edi
  loc_0055B070: mov var_40, edi
  loc_0055B073: mov var_50, edi
  loc_0055B076: mov var_60, edi
  loc_0055B079: mov var_70, edi
  loc_0055B07C: mov var_80, edi
  loc_0055B07F: mov var_90, edi
  loc_0055B085: mov var_B4, edi
  loc_0055B08B: push 00000001h
  loc_0055B08D: call [004010BCh] ; __vbaOnError
  loc_0055B093: cmp [00668024h], edi
  loc_0055B099: jnz 0055B0AFh
  loc_0055B09B: push 00668024h
  loc_0055B0A0: push 0042DEFCh
  loc_0055B0A5: mov ebx, [004011E8h] ; __vbaNew2
  loc_0055B0AB: call ebx
  loc_0055B0AD: jmp 0055B0B5h
  loc_0055B0AF: mov ebx, [004011E8h] ; __vbaNew2
  loc_0055B0B5: mov esi, [00668024h]
  loc_0055B0BB: mov eax, [esi]
  loc_0055B0BD: lea ecx, var_B4
  loc_0055B0C3: push ecx
  loc_0055B0C4: push esi
  loc_0055B0C5: call [eax+00000088h]
  loc_0055B0CB: fnclex
  loc_0055B0CD: cmp eax, edi
  loc_0055B0CF: jge 0055B0E3h
  loc_0055B0D1: push 00000088h
  loc_0055B0D6: push 0042E1B0h
  loc_0055B0DB: push esi
  loc_0055B0DC: push eax
  loc_0055B0DD: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B0E3: cmp var_B4, 00000001h
  loc_0055B0EA: jnz 0055B121h
  loc_0055B0EC: cmp [00668024h], edi
  loc_0055B0F2: jnz 0055B100h
  loc_0055B0F4: push 00668024h
  loc_0055B0F9: push 0042DEFCh
  loc_0055B0FE: call ebx
  loc_0055B100: mov esi, [00668024h]
  loc_0055B106: mov edx, [esi]
  loc_0055B108: push esi
  loc_0055B109: call [edx+0000003Ch]
  loc_0055B10C: fnclex
  loc_0055B10E: cmp eax, edi
  loc_0055B110: jge 0055B121h
  loc_0055B112: push 0000003Ch
  loc_0055B114: push 0042E1B0h
  loc_0055B119: push esi
  loc_0055B11A: push eax
  loc_0055B11B: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B121: cmp [00668024h], edi
  loc_0055B127: jnz 0055B135h
  loc_0055B129: push 00668024h
  loc_0055B12E: push 0042DEFCh
  loc_0055B133: call ebx
  loc_0055B135: mov edi, [00668024h]
  loc_0055B13B: mov var_C8, edi
  loc_0055B141: mov eax, [0066A318h]
  loc_0055B146: test eax, eax
  loc_0055B148: jnz 0055B156h
  loc_0055B14A: push 0066A318h
  loc_0055B14F: push 0042E390h
  loc_0055B154: call ebx
  loc_0055B156: mov esi, [0066A318h]
  loc_0055B15C: mov eax, [esi]
  loc_0055B15E: lea ecx, var_30
  loc_0055B161: push ecx
  loc_0055B162: push esi
  loc_0055B163: call [eax+00000014h]
  loc_0055B166: fnclex
  loc_0055B168: test eax, eax
  loc_0055B16A: jge 0055B17Bh
  loc_0055B16C: push 00000014h
  loc_0055B16E: push 0042E380h
  loc_0055B173: push esi
  loc_0055B174: push eax
  loc_0055B175: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B17B: mov eax, var_30
  loc_0055B17E: mov esi, eax
  loc_0055B180: mov edx, [eax]
  loc_0055B182: lea ecx, var_24
  loc_0055B185: push ecx
  loc_0055B186: push eax
  loc_0055B187: call [edx+00000050h]
  loc_0055B18A: fnclex
  loc_0055B18C: test eax, eax
  loc_0055B18E: jge 0055B19Fh
  loc_0055B190: push 00000050h
  loc_0055B192: push 0042E528h
  loc_0055B197: push esi
  loc_0055B198: push eax
  loc_0055B199: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B19F: mov ebx, [edi]
  loc_0055B1A1: push FFFFFFFFh
  loc_0055B1A3: push 0042DFECh
  loc_0055B1A8: push 0042DFECh
  loc_0055B1AD: push 0042E4CCh ; "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source="
  loc_0055B1B2: mov edx, var_24
  loc_0055B1B5: push edx
  loc_0055B1B6: mov esi, [00401070h] ; __vbaStrCat
  loc_0055B1BC: call __vbaStrCat
  loc_0055B1BE: mov edx, eax
  loc_0055B1C0: lea ecx, var_28
  loc_0055B1C3: mov edi, [0040126Ch] ; __vbaStrMove
  loc_0055B1C9: call edi
  loc_0055B1CB: push eax
  loc_0055B1CC: push 0042E560h ; "\settingServer.mdb;Jet OLEDB:Database Password=azis123;"
  loc_0055B1D1: call __vbaStrCat
  loc_0055B1D3: mov edx, eax
  loc_0055B1D5: lea ecx, var_2C
  loc_0055B1D8: call edi
  loc_0055B1DA: push eax
  loc_0055B1DB: mov esi, var_C8
  loc_0055B1E1: push esi
  loc_0055B1E2: call [ebx+00000050h]
  loc_0055B1E5: fnclex
  loc_0055B1E7: test eax, eax
  loc_0055B1E9: jge 0055B1FAh
  loc_0055B1EB: push 00000050h
  loc_0055B1ED: push 0042E1B0h
  loc_0055B1F2: push esi
  loc_0055B1F3: push eax
  loc_0055B1F4: call [00401080h] ; __vbaHresultCheckObj
  loc_0055B1FA: lea eax, var_2C
  loc_0055B1FD: push eax
  loc_0055B1FE: lea ecx, var_28
  loc_0055B201: push ecx
  loc_0055B202: lea edx, var_24
  loc_0055B205: push edx
  loc_0055B206: push 00000003h
  loc_0055B208: call [00401204h] ; __vbaFreeStrList
  loc_0055B20E: add esp, 00000010h
  loc_0055B211: lea ecx, var_30
  loc_0055B214: call [004012A0h] ; __vbaFreeObj
  loc_0055B21A: mov var_20, FFFFFFFFh
  loc_0055B221: call [004010A4h] ; __vbaExitProc
  loc_0055B227: push 0055B309h
  loc_0055B22C: jmp 0055B308h
  loc_0055B231: mov var_20, 00000000h
  loc_0055B238: mov ecx, 80020004h
  loc_0055B23D: mov var_68, ecx
  loc_0055B240: mov eax, 0000000Ah
  loc_0055B245: mov var_70, eax
  loc_0055B248: mov var_58, ecx
  loc_0055B24B: mov var_60, eax
  loc_0055B24E: mov var_88, 0042E624h ; "Error"
  loc_0055B258: mov edi, 00000008h
  loc_0055B25D: mov var_90, edi
  loc_0055B263: lea edx, var_90
  loc_0055B269: lea ecx, var_50
  loc_0055B26C: mov esi, [00401238h] ; __vbaVarDup
  loc_0055B272: call __vbaVarDup
  loc_0055B274: mov var_78, 0042E5D4h ; "Koneksi ke Database tidak berhasil !!"
  loc_0055B27B: mov var_80, edi
  loc_0055B27E: lea edx, var_80
  loc_0055B281: lea ecx, var_40
  loc_0055B284: call __vbaVarDup
  loc_0055B286: lea eax, var_70
  loc_0055B289: push eax
  loc_0055B28A: lea ecx, var_60
  loc_0055B28D: push ecx
  loc_0055B28E: lea edx, var_50
  loc_0055B291: push edx
  loc_0055B292: push 00000010h
  loc_0055B294: lea eax, var_40
  loc_0055B297: push eax
  loc_0055B298: call [004010B4h] ; rtcMsgBox
  loc_0055B29E: lea ecx, var_70
  loc_0055B2A1: push ecx
  loc_0055B2A2: lea edx, var_60
  loc_0055B2A5: push edx
  loc_0055B2A6: lea eax, var_50
  loc_0055B2A9: push eax
  loc_0055B2AA: lea ecx, var_40
  loc_0055B2AD: push ecx
  loc_0055B2AE: push 00000004h
  loc_0055B2B0: call [0040103Ch] ; __vbaFreeVarList
  loc_0055B2B6: add esp, 00000014h
  loc_0055B2B9: call [00401038h] ; __vbaEnd
  loc_0055B2BF: call [004010A4h] ; __vbaExitProc
  loc_0055B2C5: push 0055B309h
  loc_0055B2CA: jmp 0055B308h
  loc_0055B2CC: lea edx, var_2C
  loc_0055B2CF: push edx
  loc_0055B2D0: lea eax, var_28
  loc_0055B2D3: push eax
  loc_0055B2D4: lea ecx, var_24
  loc_0055B2D7: push ecx
  loc_0055B2D8: push 00000003h
  loc_0055B2DA: call [00401204h] ; __vbaFreeStrList
  loc_0055B2E0: add esp, 00000010h
  loc_0055B2E3: lea ecx, var_30
  loc_0055B2E6: call [004012A0h] ; __vbaFreeObj
  loc_0055B2EC: lea edx, var_70
  loc_0055B2EF: push edx
  loc_0055B2F0: lea eax, var_60
  loc_0055B2F3: push eax
  loc_0055B2F4: lea ecx, var_50
  loc_0055B2F7: push ecx
  loc_0055B2F8: lea edx, var_40
  loc_0055B2FB: push edx
  loc_0055B2FC: push 00000004h
  loc_0055B2FE: call [0040103Ch] ; __vbaFreeVarList
  loc_0055B304: add esp, 00000014h
  loc_0055B307: ret
  loc_0055B308: ret
  loc_0055B309: mov ax, var_20
  loc_0055B30D: mov ecx, var_1C
  loc_0055B310: mov fs:[00000000h], ecx
  loc_0055B317: pop edi
  loc_0055B318: pop esi
  loc_0055B319: pop ebx
  loc_0055B31A: mov esp, ebp
  loc_0055B31C: pop ebp
  loc_0055B31D: ret
  loc_0055B31E: nop
End Sub
