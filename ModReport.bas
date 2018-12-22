
Private Sub Proc_34_0_609EC0
  loc_00609EC0: push ebp
  loc_00609EC1: mov ebp, esp
  loc_00609EC3: sub esp, 00000008h
  loc_00609EC6: push 00405696h ; __vbaExceptHandler
  loc_00609ECB: mov eax, fs:[00000000h]
  loc_00609ED1: push eax
  loc_00609ED2: mov fs:[00000000h], esp
  loc_00609ED9: sub esp, 000000D0h
  loc_00609EDF: push ebx
  loc_00609EE0: push esi
  loc_00609EE1: push edi
  loc_00609EE2: mov var_8, esp
  loc_00609EE5: mov var_4, 004044A0h
  loc_00609EEC: xor ebx, ebx
  loc_00609EEE: mov var_18, ebx
  loc_00609EF1: mov var_1C, ebx
  loc_00609EF4: mov var_20, ebx
  loc_00609EF7: mov var_24, ebx
  loc_00609EFA: mov var_28, ebx
  loc_00609EFD: mov var_2C, ebx
  loc_00609F00: mov var_30, ebx
  loc_00609F03: mov var_40, ebx
  loc_00609F06: mov var_50, ebx
  loc_00609F09: mov var_60, ebx
  loc_00609F0C: mov var_70, ebx
  loc_00609F0F: mov var_80, ebx
  loc_00609F12: mov var_90, ebx
  loc_00609F18: mov var_A4, ebx
  loc_00609F1E: mov var_C4, ebx
  loc_00609F24: mov var_C8, ebx
  loc_00609F2A: call 0055B320h
  loc_00609F2F: mov esi, [004011FCh] ; __vbaStrCopy
  loc_00609F35: mov edx, 0042DFECh
  loc_00609F3A: lea ecx, var_18
  loc_00609F3D: call __vbaStrCopy
  loc_00609F3F: mov edx, 0043B6D8h ; " SELECT userid,nama,level FROM userpengguna ORDER BY userid"
  loc_00609F44: lea ecx, var_18
  loc_00609F47: call __vbaStrCopy
  loc_00609F49: push 0042DF30h
  loc_00609F4E: call [00401168h] ; __vbaNew
  loc_00609F54: push eax
  loc_00609F55: push 00668090h
  loc_00609F5A: call [004010B8h] ; __vbaObjSet
  loc_00609F60: cmp [0066803Ch], ebx
  loc_00609F66: jnz 00609F78h
  loc_00609F68: push 0066803Ch
  loc_00609F6D: push 0042DEFCh
  loc_00609F72: call [004011E8h] ; __vbaNew2
  loc_00609F78: mov esi, [0066803Ch]
  loc_00609F7E: lea ecx, var_40
  loc_00609F81: mov var_38, 80020004h
  loc_00609F88: mov var_40, 0000000Ah
  loc_00609F8F: call [0040123Ch] ; __vbaFreeVarg
  loc_00609F95: mov eax, [esi]
  loc_00609F97: lea ecx, var_24
  loc_00609F9A: push ecx
  loc_00609F9B: mov ecx, var_18
  loc_00609F9E: lea edx, var_40
  loc_00609FA1: push 00000001h
  loc_00609FA3: push edx
  loc_00609FA4: push ecx
  loc_00609FA5: push esi
  loc_00609FA6: call [eax+00000040h]
  loc_00609FA9: cmp eax, ebx
  loc_00609FAB: fnclex
  loc_00609FAD: jge 00609FC2h
  loc_00609FAF: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00609FB5: push 00000040h
  loc_00609FB7: push 0042E1B0h
  loc_00609FBC: push esi
  loc_00609FBD: push eax
  loc_00609FBE: call edi
  loc_00609FC0: jmp 00609FC8h
  loc_00609FC2: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00609FC8: mov edx, var_24
  loc_00609FCB: mov esi, [00401268h] ; __vbaCastObj
  loc_00609FD1: push 0042DF20h
  loc_00609FD6: push edx
  loc_00609FD7: call __vbaCastObj
  loc_00609FD9: push eax
  loc_00609FDA: push 00668090h
  loc_00609FDF: call [004010B8h] ; __vbaObjSet
  loc_00609FE5: lea ecx, var_24
  loc_00609FE8: call [004012A0h] ; __vbaFreeObj
  loc_00609FEE: lea ecx, var_40
  loc_00609FF1: call [00401020h] ; __vbaFreeVar
  loc_00609FF7: cmp [00668508h], ebx
  loc_00609FFD: jnz 0060A00Fh
  loc_00609FFF: push 00668508h
  loc_0060A004: push 00407328h
  loc_0060A009: call [004011E8h] ; __vbaNew2
  loc_0060A00F: mov eax, [00668508h]
  loc_0060A014: lea ecx, var_C4
  loc_0060A01A: push eax
  loc_0060A01B: push ecx
  loc_0060A01C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060A022: mov edx, var_C4
  loc_0060A028: push 0043B7B8h
  loc_0060A02D: push ebx
  loc_0060A02E: mov edx, [edx]
  loc_0060A030: mov var_E0, edx
  loc_0060A036: call __vbaCastObj
  loc_0060A038: push eax
  loc_0060A039: lea eax, var_24
  loc_0060A03C: push eax
  loc_0060A03D: call [004010B8h] ; __vbaObjSet
  loc_0060A043: mov ecx, var_C4
  loc_0060A049: mov edx, var_E0
  loc_0060A04F: push eax
  loc_0060A050: push ecx
  loc_0060A051: call [edx+00000028h]
  loc_0060A054: cmp eax, ebx
  loc_0060A056: fnclex
  loc_0060A058: jge 0060A06Bh
  loc_0060A05A: mov ecx, var_C4
  loc_0060A060: push 00000028h
  loc_0060A062: push 0043B5D0h
  loc_0060A067: push ecx
  loc_0060A068: push eax
  loc_0060A069: call edi
  loc_0060A06B: lea ecx, var_24
  loc_0060A06E: call [004012A0h] ; __vbaFreeObj
  loc_0060A074: mov eax, var_C4
  loc_0060A07A: push 0042DFECh
  loc_0060A07F: push eax
  loc_0060A080: mov edx, [eax]
  loc_0060A082: call [edx+00000030h]
  loc_0060A085: cmp eax, ebx
  loc_0060A087: fnclex
  loc_0060A089: jge 0060A09Ch
  loc_0060A08B: mov ecx, var_C4
  loc_0060A091: push 00000030h
  loc_0060A093: push 0043B5D0h
  loc_0060A098: push ecx
  loc_0060A099: push eax
  loc_0060A09A: call edi
  loc_0060A09C: mov eax, [00668090h]
  loc_0060A0A1: lea ecx, var_24
  loc_0060A0A4: push ecx
  loc_0060A0A5: push eax
  loc_0060A0A6: mov edx, [eax]
  loc_0060A0A8: call [edx+00000114h]
  loc_0060A0AE: cmp eax, ebx
  loc_0060A0B0: fnclex
  loc_0060A0B2: jge 0060A0C8h
  loc_0060A0B4: mov edx, [00668090h]
  loc_0060A0BA: push 00000114h
  loc_0060A0BF: push 0043B7C8h
  loc_0060A0C4: push edx
  loc_0060A0C5: push eax
  loc_0060A0C6: call edi
  loc_0060A0C8: mov eax, var_C4
  loc_0060A0CE: mov ecx, var_24
  loc_0060A0D1: push 0043B7B8h
  loc_0060A0D6: push ecx
  loc_0060A0D7: mov esi, [eax]
  loc_0060A0D9: call [00401268h] ; __vbaCastObj
  loc_0060A0DF: lea edx, var_28
  loc_0060A0E2: push eax
  loc_0060A0E3: push edx
  loc_0060A0E4: call [004010B8h] ; __vbaObjSet
  loc_0060A0EA: push eax
  loc_0060A0EB: mov eax, var_C4
  loc_0060A0F1: push eax
  loc_0060A0F2: call [esi+00000028h]
  loc_0060A0F5: cmp eax, ebx
  loc_0060A0F7: fnclex
  loc_0060A0F9: jge 0060A10Ch
  loc_0060A0FB: mov ecx, var_C4
  loc_0060A101: push 00000028h
  loc_0060A103: push 0043B5D0h
  loc_0060A108: push ecx
  loc_0060A109: push eax
  loc_0060A10A: call edi
  loc_0060A10C: lea edx, var_28
  loc_0060A10F: lea eax, var_24
  loc_0060A112: push edx
  loc_0060A113: push eax
  loc_0060A114: push 00000002h
  loc_0060A116: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A11C: mov eax, var_C4
  loc_0060A122: add esp, 0000000Ch
  loc_0060A125: lea edx, var_24
  loc_0060A128: mov ecx, [eax]
  loc_0060A12A: push edx
  loc_0060A12B: push eax
  loc_0060A12C: call [ecx+0000008Ch]
  loc_0060A132: cmp eax, ebx
  loc_0060A134: fnclex
  loc_0060A136: jge 0060A14Ch
  loc_0060A138: mov ecx, var_C4
  loc_0060A13E: push 0000008Ch
  loc_0060A143: push 0043B5D0h
  loc_0060A148: push ecx
  loc_0060A149: push eax
  loc_0060A14A: call edi
  loc_0060A14C: lea esi, var_28
  loc_0060A14F: mov eax, var_24
  loc_0060A152: push esi
  loc_0060A153: mov ecx, 00000008h
  loc_0060A158: sub esp, 00000010h
  loc_0060A15B: mov var_70, ecx
  loc_0060A15E: mov esi, esp
  loc_0060A160: mov var_68, 00439FB0h ; "Section1"
  loc_0060A167: mov edx, [eax]
  loc_0060A169: push eax
  loc_0060A16A: mov [esi], ecx
  loc_0060A16C: mov ecx, var_6C
  loc_0060A16F: mov var_AC, eax
  loc_0060A175: mov [esi+00000004h], ecx
  loc_0060A178: mov ecx, var_68
  loc_0060A17B: mov [esi+00000008h], ecx
  loc_0060A17E: mov ecx, var_64
  loc_0060A181: mov [esi+0000000Ch], ecx
  loc_0060A184: call [edx+0000001Ch]
  loc_0060A187: cmp eax, ebx
  loc_0060A189: fnclex
  loc_0060A18B: jge 0060A19Eh
  loc_0060A18D: mov edx, var_AC
  loc_0060A193: push 0000001Ch
  loc_0060A195: push 00439DD8h
  loc_0060A19A: push edx
  loc_0060A19B: push eax
  loc_0060A19C: call edi
  loc_0060A19E: mov eax, var_28
  loc_0060A1A1: lea edx, var_2C
  loc_0060A1A4: push edx
  loc_0060A1A5: push eax
  loc_0060A1A6: mov ecx, [eax]
  loc_0060A1A8: mov esi, eax
  loc_0060A1AA: call [ecx+0000003Ch]
  loc_0060A1AD: cmp eax, ebx
  loc_0060A1AF: fnclex
  loc_0060A1B1: jge 0060A1BEh
  loc_0060A1B3: push 0000003Ch
  loc_0060A1B5: push 00439BF8h
  loc_0060A1BA: push esi
  loc_0060A1BB: push eax
  loc_0060A1BC: call edi
  loc_0060A1BE: mov eax, var_2C
  loc_0060A1C1: mov var_2C, ebx
  loc_0060A1C4: push eax
  loc_0060A1C5: lea eax, var_C8
  loc_0060A1CB: push eax
  loc_0060A1CC: call [004010B8h] ; __vbaObjSet
  loc_0060A1D2: lea ecx, var_28
  loc_0060A1D5: lea edx, var_24
  loc_0060A1D8: push ecx
  loc_0060A1D9: push edx
  loc_0060A1DA: push 00000002h
  loc_0060A1DC: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A1E2: mov eax, var_C8
  loc_0060A1E8: add esp, 0000000Ch
  loc_0060A1EB: lea edx, var_A4
  loc_0060A1F1: mov ecx, [eax]
  loc_0060A1F3: push edx
  loc_0060A1F4: push eax
  loc_0060A1F5: call [ecx+00000020h]
  loc_0060A1F8: cmp eax, ebx
  loc_0060A1FA: fnclex
  loc_0060A1FC: jge 0060A20Fh
  loc_0060A1FE: mov ecx, var_C8
  loc_0060A204: push 00000020h
  loc_0060A206: push 0043B7D8h
  loc_0060A20B: push ecx
  loc_0060A20C: push eax
  loc_0060A20D: call edi
  loc_0060A20F: mov ecx, var_A4
  loc_0060A215: call [00401138h] ; __vbaI2I4
  loc_0060A21B: mov var_D0, eax
  loc_0060A221: mov ebx, 00000001h
  loc_0060A226: cmp bx, var_D0
  loc_0060A22D: mov var_14, ebx
  loc_0060A230: jg 0060A4A1h
  loc_0060A236: lea esi, var_24
  loc_0060A239: mov ecx, var_C8
  loc_0060A23F: push esi
  loc_0060A240: mov eax, 00000002h
  loc_0060A245: sub esp, 00000010h
  loc_0060A248: mov var_70, eax
  loc_0060A24B: mov esi, esp
  loc_0060A24D: mov var_68, bx
  loc_0060A251: mov edx, [ecx]
  loc_0060A253: push ecx
  loc_0060A254: mov [esi], eax
  loc_0060A256: mov eax, var_6C
  loc_0060A259: mov [esi+00000004h], eax
  loc_0060A25C: mov eax, var_68
  loc_0060A25F: mov [esi+00000008h], eax
  loc_0060A262: mov eax, var_64
  loc_0060A265: mov [esi+0000000Ch], eax
  loc_0060A268: call [edx+0000001Ch]
  loc_0060A26B: test eax, eax
  loc_0060A26D: fnclex
  loc_0060A26F: jge 0060A282h
  loc_0060A271: mov ecx, var_C8
  loc_0060A277: push 0000001Ch
  loc_0060A279: push 0043B7D8h
  loc_0060A27E: push ecx
  loc_0060A27F: push eax
  loc_0060A280: call edi
  loc_0060A282: mov edx, var_24
  loc_0060A285: push 0043B7E8h
  loc_0060A28A: push edx
  loc_0060A28B: call [004011CCh] ; __vbaCheckType
  loc_0060A291: lea ecx, var_24
  loc_0060A294: mov si, ax
  loc_0060A297: call [004012A0h] ; __vbaFreeObj
  loc_0060A29D: test si, si
  loc_0060A2A0: Unknown_795()
  loc_0060A2A6: mov var_68, bx
  loc_0060A2AA: lea ebx, var_24
  loc_0060A2AD: push ebx
  loc_0060A2AE: mov ecx, var_C8
  loc_0060A2B4: sub esp, 00000010h
  loc_0060A2B7: mov eax, 00000002h
  loc_0060A2BC: mov ebx, esp
  loc_0060A2BE: mov var_70, eax
  loc_0060A2C1: mov edx, [ecx]
  loc_0060A2C3: mov esi, 0042DFECh
  loc_0060A2C8: mov [ebx], eax
  loc_0060A2CA: mov eax, var_6C
  loc_0060A2CD: push ecx
  loc_0060A2CE: mov var_78, esi
  loc_0060A2D1: mov [ebx+00000004h], eax
  loc_0060A2D4: mov eax, var_68
  loc_0060A2D7: mov edi, 00000008h
  loc_0060A2DC: mov [ebx+00000008h], eax
  loc_0060A2DF: mov eax, var_64
  loc_0060A2E2: mov [ebx+0000000Ch], eax
  loc_0060A2E5: call [edx+0000001Ch]
  loc_0060A2E8: test eax, eax
  loc_0060A2EA: fnclex
  loc_0060A2EC: jge 0060A303h
  loc_0060A2EE: mov ecx, var_C8
  loc_0060A2F4: push 0000001Ch
  loc_0060A2F6: push 0043B7D8h
  loc_0060A2FB: push ecx
  loc_0060A2FC: push eax
  loc_0060A2FD: call [00401080h] ; __vbaHresultCheckObj
  loc_0060A303: mov eax, var_7C
  loc_0060A306: sub esp, 00000010h
  loc_0060A309: mov edx, esp
  loc_0060A30B: mov ecx, var_74
  loc_0060A30E: push 0043B7F8h ; "DataMember"
  loc_0060A313: mov [edx], edi
  loc_0060A315: mov [edx+00000004h], eax
  loc_0060A318: mov [edx+00000008h], esi
  loc_0060A31B: mov [edx+0000000Ch], ecx
  loc_0060A31E: mov edx, var_24
  loc_0060A321: push edx
  loc_0060A322: call [00401094h] ; __vbaLateMemSt
  loc_0060A328: lea ecx, var_24
  loc_0060A32B: call [004012A0h] ; __vbaFreeObj
  loc_0060A331: mov eax, [00668090h]
  loc_0060A336: lea edx, var_24
  loc_0060A339: push edx
  loc_0060A33A: push eax
  loc_0060A33B: mov ecx, [eax]
  loc_0060A33D: call [ecx+00000054h]
  loc_0060A340: test eax, eax
  loc_0060A342: fnclex
  loc_0060A344: jge 0060A35Bh
  loc_0060A346: mov ecx, [00668090h]
  loc_0060A34C: push 00000054h
  loc_0060A34E: push 0042DF44h
  loc_0060A353: push ecx
  loc_0060A354: push eax
  loc_0060A355: call [00401080h] ; __vbaHresultCheckObj
  loc_0060A35B: mov ebx, var_14
  loc_0060A35E: lea edi, var_28
  loc_0060A361: mov dx, bx
  loc_0060A364: push edi
  loc_0060A365: sub dx, 0001h
  loc_0060A369: mov eax, var_24
  loc_0060A36C: jo 0060AC2Fh
  loc_0060A372: sub esp, 00000010h
  loc_0060A375: mov ecx, 00000002h
  loc_0060A37A: mov edi, esp
  loc_0060A37C: mov var_70, ecx
  loc_0060A37F: mov var_68, dx
  loc_0060A383: mov edx, [eax]
  loc_0060A385: mov [edi], ecx
  loc_0060A387: mov ecx, var_6C
  loc_0060A38A: push eax
  loc_0060A38B: mov esi, eax
  loc_0060A38D: mov [edi+00000004h], ecx
  loc_0060A390: mov ecx, var_68
  loc_0060A393: mov [edi+00000008h], ecx
  loc_0060A396: mov ecx, var_64
  loc_0060A399: mov [edi+0000000Ch], ecx
  loc_0060A39C: call [edx+00000028h]
  loc_0060A39F: test eax, eax
  loc_0060A3A1: fnclex
  loc_0060A3A3: jge 0060A3B8h
  loc_0060A3A5: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060A3AB: push 00000028h
  loc_0060A3AD: push 0042DFACh
  loc_0060A3B2: push esi
  loc_0060A3B3: push eax
  loc_0060A3B4: call edi
  loc_0060A3B6: jmp 0060A3BEh
  loc_0060A3B8: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060A3BE: mov eax, var_28
  loc_0060A3C1: lea ecx, var_20
  loc_0060A3C4: push ecx
  loc_0060A3C5: push eax
  loc_0060A3C6: mov edx, [eax]
  loc_0060A3C8: mov esi, eax
  loc_0060A3CA: call [edx+0000002Ch]
  loc_0060A3CD: test eax, eax
  loc_0060A3CF: fnclex
  loc_0060A3D1: jge 0060A3DEh
  loc_0060A3D3: push 0000002Ch
  loc_0060A3D5: push 0042DFBCh
  loc_0060A3DA: push esi
  loc_0060A3DB: push eax
  loc_0060A3DC: call edi
  loc_0060A3DE: mov eax, var_20
  loc_0060A3E1: lea esi, var_2C
  loc_0060A3E4: push esi
  loc_0060A3E5: mov ecx, var_C8
  loc_0060A3EB: sub esp, 00000010h
  loc_0060A3EE: mov var_38, eax
  loc_0060A3F1: mov esi, esp
  loc_0060A3F3: mov eax, 00000002h
  loc_0060A3F8: mov var_78, bx
  loc_0060A3FC: mov var_20, 00000000h
  loc_0060A403: mov [esi], eax
  loc_0060A405: mov eax, var_7C
  loc_0060A408: mov var_40, 00000008h
  loc_0060A40F: mov edx, [ecx]
  loc_0060A411: mov [esi+00000004h], eax
  loc_0060A414: mov eax, var_78
  loc_0060A417: push ecx
  loc_0060A418: mov [esi+00000008h], eax
  loc_0060A41B: mov eax, var_74
  loc_0060A41E: mov [esi+0000000Ch], eax
  loc_0060A421: call [edx+0000001Ch]
  loc_0060A424: test eax, eax
  loc_0060A426: fnclex
  loc_0060A428: jge 0060A43Bh
  loc_0060A42A: mov ecx, var_C8
  loc_0060A430: push 0000001Ch
  loc_0060A432: push 0043B7D8h
  loc_0060A437: push ecx
  loc_0060A438: push eax
  loc_0060A439: call edi
  loc_0060A43B: mov eax, var_40
  loc_0060A43E: mov ecx, var_3C
  loc_0060A441: sub esp, 00000010h
  loc_0060A444: mov edx, esp
  loc_0060A446: push 0043B810h ; "DataField"
  loc_0060A44B: mov [edx], eax
  loc_0060A44D: mov eax, var_38
  loc_0060A450: mov [edx+00000004h], ecx
  loc_0060A453: mov ecx, var_34
  loc_0060A456: mov [edx+00000008h], eax
  loc_0060A459: mov [edx+0000000Ch], ecx
  loc_0060A45C: mov edx, var_2C
  loc_0060A45F: push edx
  loc_0060A460: call [00401094h] ; __vbaLateMemSt
  loc_0060A466: lea eax, var_2C
  loc_0060A469: lea ecx, var_28
  loc_0060A46C: push eax
  loc_0060A46D: lea edx, var_24
  loc_0060A470: push ecx
  loc_0060A471: push edx
  loc_0060A472: push 00000003h
  loc_0060A474: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A47A: add esp, 00000010h
  loc_0060A47D: lea ecx, var_40
  loc_0060A480: call [00401020h] ; __vbaFreeVar
  loc_0060A486: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060A48C: mov eax, 00000001h
  loc_0060A491: add ax, bx
  loc_0060A494: jo 0060AC2Fh
  loc_0060A49A: mov ebx, eax
  loc_0060A49C: jmp 0060A226h
  loc_0060A4A1: lea eax, var_C8
  loc_0060A4A7: push 00000000h
  loc_0060A4A9: push eax
  loc_0060A4AA: call [004010C8h] ; __vbaObjSetAddref
  loc_0060A4B0: lea ecx, var_40
  loc_0060A4B3: push ecx
  loc_0060A4B4: call [00401220h] ; rtcGetDateVar
  loc_0060A4BA: lea edx, var_70
  loc_0060A4BD: lea ecx, var_50
  loc_0060A4C0: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060A4C7: mov var_70, 00000008h
  loc_0060A4CE: call [00401238h] ; __vbaVarDup
  loc_0060A4D4: push 00000001h
  loc_0060A4D6: lea edx, var_50
  loc_0060A4D9: push 00000001h
  loc_0060A4DB: lea eax, var_40
  loc_0060A4DE: push edx
  loc_0060A4DF: lea ecx, var_60
  loc_0060A4E2: push eax
  loc_0060A4E3: push ecx
  loc_0060A4E4: call [00401078h] ; rtcVarFromFormatVar
  loc_0060A4EA: mov eax, var_C4
  loc_0060A4F0: lea ecx, var_24
  loc_0060A4F3: push ecx
  loc_0060A4F4: push eax
  loc_0060A4F5: mov edx, [eax]
  loc_0060A4F7: call [edx+0000008Ch]
  loc_0060A4FD: test eax, eax
  loc_0060A4FF: fnclex
  loc_0060A501: jge 0060A517h
  loc_0060A503: mov edx, var_C4
  loc_0060A509: push 0000008Ch
  loc_0060A50E: push 0043B5D0h
  loc_0060A513: push edx
  loc_0060A514: push eax
  loc_0060A515: call edi
  loc_0060A517: lea ebx, var_28
  loc_0060A51A: mov eax, var_24
  loc_0060A51D: push ebx
  loc_0060A51E: mov edx, 00000008h
  loc_0060A523: sub esp, 00000010h
  loc_0060A526: mov esi, [eax]
  loc_0060A528: mov ebx, esp
  loc_0060A52A: mov ecx, 0043B848h ; "Section4"
  loc_0060A52F: push eax
  loc_0060A530: mov var_AC, eax
  loc_0060A536: mov [ebx], edx
  loc_0060A538: mov edx, var_7C
  loc_0060A53B: mov [ebx+00000004h], edx
  loc_0060A53E: mov [ebx+00000008h], ecx
  loc_0060A541: mov ecx, var_74
  loc_0060A544: mov [ebx+0000000Ch], ecx
  loc_0060A547: call [esi+0000001Ch]
  loc_0060A54A: test eax, eax
  loc_0060A54C: fnclex
  loc_0060A54E: jge 0060A561h
  loc_0060A550: mov edx, var_AC
  loc_0060A556: push 0000001Ch
  loc_0060A558: push 00439DD8h
  loc_0060A55D: push edx
  loc_0060A55E: push eax
  loc_0060A55F: call edi
  loc_0060A561: mov eax, var_28
  loc_0060A564: lea edx, var_2C
  loc_0060A567: push edx
  loc_0060A568: push eax
  loc_0060A569: mov ecx, [eax]
  loc_0060A56B: mov esi, eax
  loc_0060A56D: call [ecx+0000003Ch]
  loc_0060A570: test eax, eax
  loc_0060A572: fnclex
  loc_0060A574: jge 0060A581h
  loc_0060A576: push 0000003Ch
  loc_0060A578: push 00439BF8h
  loc_0060A57D: push esi
  loc_0060A57E: push eax
  loc_0060A57F: call edi
  loc_0060A581: lea ebx, var_30
  loc_0060A584: mov eax, var_2C
  loc_0060A587: push ebx
  loc_0060A588: mov edx, 00000008h
  loc_0060A58D: sub esp, 00000010h
  loc_0060A590: mov esi, [eax]
  loc_0060A592: mov ebx, esp
  loc_0060A594: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060A599: push eax
  loc_0060A59A: mov var_BC, eax
  loc_0060A5A0: mov [ebx], edx
  loc_0060A5A2: mov edx, var_8C
  loc_0060A5A8: mov [ebx+00000004h], edx
  loc_0060A5AB: mov [ebx+00000008h], ecx
  loc_0060A5AE: mov ecx, var_84
  loc_0060A5B4: mov [ebx+0000000Ch], ecx
  loc_0060A5B7: call [esi+0000001Ch]
  loc_0060A5BA: test eax, eax
  loc_0060A5BC: fnclex
  loc_0060A5BE: jge 0060A5D1h
  loc_0060A5C0: mov edx, var_BC
  loc_0060A5C6: push 0000001Ch
  loc_0060A5C8: push 0043B7D8h
  loc_0060A5CD: push edx
  loc_0060A5CE: push eax
  loc_0060A5CF: call edi
  loc_0060A5D1: mov ecx, var_60
  loc_0060A5D4: mov edx, var_5C
  loc_0060A5D7: sub esp, 00000010h
  loc_0060A5DA: mov eax, esp
  loc_0060A5DC: push 0043B85Ch ; "Caption"
  loc_0060A5E1: mov [eax], ecx
  loc_0060A5E3: mov ecx, var_58
  loc_0060A5E6: mov [eax+00000004h], edx
  loc_0060A5E9: mov edx, var_54
  loc_0060A5EC: mov [eax+00000008h], ecx
  loc_0060A5EF: mov [eax+0000000Ch], edx
  loc_0060A5F2: mov eax, var_30
  loc_0060A5F5: push eax
  loc_0060A5F6: call [00401094h] ; __vbaLateMemSt
  loc_0060A5FC: lea ecx, var_30
  loc_0060A5FF: lea edx, var_2C
  loc_0060A602: push ecx
  loc_0060A603: lea eax, var_28
  loc_0060A606: push edx
  loc_0060A607: lea ecx, var_24
  loc_0060A60A: push eax
  loc_0060A60B: push ecx
  loc_0060A60C: push 00000004h
  loc_0060A60E: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A614: lea edx, var_60
  loc_0060A617: lea eax, var_50
  loc_0060A61A: push edx
  loc_0060A61B: lea ecx, var_40
  loc_0060A61E: push eax
  loc_0060A61F: push ecx
  loc_0060A620: push 00000003h
  loc_0060A622: call [0040103Ch] ; __vbaFreeVarList
  loc_0060A628: mov eax, var_C4
  loc_0060A62E: add esp, 00000024h
  loc_0060A631: lea ecx, var_24
  loc_0060A634: mov edx, [eax]
  loc_0060A636: push ecx
  loc_0060A637: push eax
  loc_0060A638: call [edx+0000008Ch]
  loc_0060A63E: test eax, eax
  loc_0060A640: fnclex
  loc_0060A642: jge 0060A658h
  loc_0060A644: mov edx, var_C4
  loc_0060A64A: push 0000008Ch
  loc_0060A64F: push 0043B5D0h
  loc_0060A654: push edx
  loc_0060A655: push eax
  loc_0060A656: call edi
  loc_0060A658: lea ebx, var_28
  loc_0060A65B: mov eax, var_24
  loc_0060A65E: push ebx
  loc_0060A65F: mov edx, 00000008h
  loc_0060A664: sub esp, 00000010h
  loc_0060A667: mov var_70, edx
  loc_0060A66A: mov ebx, esp
  loc_0060A66C: mov ecx, 0043B848h ; "Section4"
  loc_0060A671: mov var_68, ecx
  loc_0060A674: mov esi, [eax]
  loc_0060A676: mov [ebx], edx
  loc_0060A678: mov edx, var_6C
  loc_0060A67B: push eax
  loc_0060A67C: mov var_AC, eax
  loc_0060A682: mov [ebx+00000004h], edx
  loc_0060A685: mov [ebx+00000008h], ecx
  loc_0060A688: mov ecx, var_64
  loc_0060A68B: mov [ebx+0000000Ch], ecx
  loc_0060A68E: call [esi+0000001Ch]
  loc_0060A691: test eax, eax
  loc_0060A693: fnclex
  loc_0060A695: jge 0060A6A8h
  loc_0060A697: mov edx, var_AC
  loc_0060A69D: push 0000001Ch
  loc_0060A69F: push 00439DD8h
  loc_0060A6A4: push edx
  loc_0060A6A5: push eax
  loc_0060A6A6: call edi
  loc_0060A6A8: mov eax, var_28
  loc_0060A6AB: lea edx, var_2C
  loc_0060A6AE: push edx
  loc_0060A6AF: push eax
  loc_0060A6B0: mov ecx, [eax]
  loc_0060A6B2: mov esi, eax
  loc_0060A6B4: call [ecx+0000003Ch]
  loc_0060A6B7: test eax, eax
  loc_0060A6B9: fnclex
  loc_0060A6BB: jge 0060A6C8h
  loc_0060A6BD: push 0000003Ch
  loc_0060A6BF: push 00439BF8h
  loc_0060A6C4: push esi
  loc_0060A6C5: push eax
  loc_0060A6C6: call edi
  loc_0060A6C8: lea ebx, var_30
  loc_0060A6CB: mov eax, var_2C
  loc_0060A6CE: push ebx
  loc_0060A6CF: mov edx, 00000008h
  loc_0060A6D4: sub esp, 00000010h
  loc_0060A6D7: mov esi, [eax]
  loc_0060A6D9: mov ebx, esp
  loc_0060A6DB: mov ecx, 0043B870h ; "LabelNama"
  loc_0060A6E0: push eax
  loc_0060A6E1: mov var_BC, eax
  loc_0060A6E7: mov [ebx], edx
  loc_0060A6E9: mov edx, var_7C
  loc_0060A6EC: mov [ebx+00000004h], edx
  loc_0060A6EF: mov [ebx+00000008h], ecx
  loc_0060A6F2: mov ecx, var_74
  loc_0060A6F5: mov [ebx+0000000Ch], ecx
  loc_0060A6F8: call [esi+0000001Ch]
  loc_0060A6FB: test eax, eax
  loc_0060A6FD: fnclex
  loc_0060A6FF: jge 0060A712h
  loc_0060A701: mov edx, var_BC
  loc_0060A707: push 0000001Ch
  loc_0060A709: push 0043B7D8h
  loc_0060A70E: push edx
  loc_0060A70F: push eax
  loc_0060A710: call edi
  loc_0060A712: mov ecx, [006680D8h]
  loc_0060A718: mov edx, [006680DCh]
  loc_0060A71E: sub esp, 00000010h
  loc_0060A721: mov eax, esp
  loc_0060A723: push 0043B85Ch ; "Caption"
  loc_0060A728: mov [eax], ecx
  loc_0060A72A: mov ecx, [006680E0h]
  loc_0060A730: mov [eax+00000004h], edx
  loc_0060A733: mov edx, [006680E4h]
  loc_0060A739: mov [eax+00000008h], ecx
  loc_0060A73C: mov [eax+0000000Ch], edx
  loc_0060A73F: mov eax, var_30
  loc_0060A742: push eax
  loc_0060A743: call [00401094h] ; __vbaLateMemSt
  loc_0060A749: lea ecx, var_30
  loc_0060A74C: lea edx, var_2C
  loc_0060A74F: push ecx
  loc_0060A750: lea eax, var_28
  loc_0060A753: push edx
  loc_0060A754: lea ecx, var_24
  loc_0060A757: push eax
  loc_0060A758: push ecx
  loc_0060A759: push 00000004h
  loc_0060A75B: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A761: mov eax, var_C4
  loc_0060A767: add esp, 00000014h
  loc_0060A76A: lea ecx, var_24
  loc_0060A76D: mov edx, [eax]
  loc_0060A76F: push ecx
  loc_0060A770: push eax
  loc_0060A771: call [edx+0000008Ch]
  loc_0060A777: test eax, eax
  loc_0060A779: fnclex
  loc_0060A77B: jge 0060A791h
  loc_0060A77D: mov edx, var_C4
  loc_0060A783: push 0000008Ch
  loc_0060A788: push 0043B5D0h
  loc_0060A78D: push edx
  loc_0060A78E: push eax
  loc_0060A78F: call edi
  loc_0060A791: lea ebx, var_28
  loc_0060A794: mov eax, var_24
  loc_0060A797: push ebx
  loc_0060A798: mov edx, 00000008h
  loc_0060A79D: sub esp, 00000010h
  loc_0060A7A0: mov var_70, edx
  loc_0060A7A3: mov ebx, esp
  loc_0060A7A5: mov ecx, 0043B848h ; "Section4"
  loc_0060A7AA: mov var_68, ecx
  loc_0060A7AD: mov esi, [eax]
  loc_0060A7AF: mov [ebx], edx
  loc_0060A7B1: mov edx, var_6C
  loc_0060A7B4: push eax
  loc_0060A7B5: mov var_AC, eax
  loc_0060A7BB: mov [ebx+00000004h], edx
  loc_0060A7BE: mov [ebx+00000008h], ecx
  loc_0060A7C1: mov ecx, var_64
  loc_0060A7C4: mov [ebx+0000000Ch], ecx
  loc_0060A7C7: call [esi+0000001Ch]
  loc_0060A7CA: test eax, eax
  loc_0060A7CC: fnclex
  loc_0060A7CE: jge 0060A7E1h
  loc_0060A7D0: mov edx, var_AC
  loc_0060A7D6: push 0000001Ch
  loc_0060A7D8: push 00439DD8h
  loc_0060A7DD: push edx
  loc_0060A7DE: push eax
  loc_0060A7DF: call edi
  loc_0060A7E1: mov eax, var_28
  loc_0060A7E4: lea edx, var_2C
  loc_0060A7E7: push edx
  loc_0060A7E8: push eax
  loc_0060A7E9: mov ecx, [eax]
  loc_0060A7EB: mov esi, eax
  loc_0060A7ED: call [ecx+0000003Ch]
  loc_0060A7F0: test eax, eax
  loc_0060A7F2: fnclex
  loc_0060A7F4: jge 0060A801h
  loc_0060A7F6: push 0000003Ch
  loc_0060A7F8: push 00439BF8h
  loc_0060A7FD: push esi
  loc_0060A7FE: push eax
  loc_0060A7FF: call edi
  loc_0060A801: lea ebx, var_30
  loc_0060A804: mov eax, var_2C
  loc_0060A807: push ebx
  loc_0060A808: mov edx, 00000008h
  loc_0060A80D: sub esp, 00000010h
  loc_0060A810: mov esi, [eax]
  loc_0060A812: mov ebx, esp
  loc_0060A814: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060A819: push eax
  loc_0060A81A: mov var_BC, eax
  loc_0060A820: mov [ebx], edx
  loc_0060A822: mov edx, var_7C
  loc_0060A825: mov [ebx+00000004h], edx
  loc_0060A828: mov [ebx+00000008h], ecx
  loc_0060A82B: mov ecx, var_74
  loc_0060A82E: mov [ebx+0000000Ch], ecx
  loc_0060A831: call [esi+0000001Ch]
  loc_0060A834: test eax, eax
  loc_0060A836: fnclex
  loc_0060A838: jge 0060A84Bh
  loc_0060A83A: mov edx, var_BC
  loc_0060A840: push 0000001Ch
  loc_0060A842: push 0043B7D8h
  loc_0060A847: push edx
  loc_0060A848: push eax
  loc_0060A849: call edi
  loc_0060A84B: mov ecx, [006680E8h]
  loc_0060A851: mov edx, [006680ECh]
  loc_0060A857: sub esp, 00000010h
  loc_0060A85A: mov eax, esp
  loc_0060A85C: push 0043B85Ch ; "Caption"
  loc_0060A861: mov [eax], ecx
  loc_0060A863: mov ecx, [006680F0h]
  loc_0060A869: mov [eax+00000004h], edx
  loc_0060A86C: mov edx, [006680F4h]
  loc_0060A872: mov [eax+00000008h], ecx
  loc_0060A875: mov [eax+0000000Ch], edx
  loc_0060A878: mov eax, var_30
  loc_0060A87B: push eax
  loc_0060A87C: call [00401094h] ; __vbaLateMemSt
  loc_0060A882: lea ecx, var_30
  loc_0060A885: lea edx, var_2C
  loc_0060A888: push ecx
  loc_0060A889: lea eax, var_28
  loc_0060A88C: push edx
  loc_0060A88D: lea ecx, var_24
  loc_0060A890: push eax
  loc_0060A891: push ecx
  loc_0060A892: push 00000004h
  loc_0060A894: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A89A: mov eax, var_C4
  loc_0060A8A0: add esp, 00000014h
  loc_0060A8A3: lea ecx, var_24
  loc_0060A8A6: mov edx, [eax]
  loc_0060A8A8: push ecx
  loc_0060A8A9: push eax
  loc_0060A8AA: call [edx+0000008Ch]
  loc_0060A8B0: test eax, eax
  loc_0060A8B2: fnclex
  loc_0060A8B4: jge 0060A8CAh
  loc_0060A8B6: mov edx, var_C4
  loc_0060A8BC: push 0000008Ch
  loc_0060A8C1: push 0043B5D0h
  loc_0060A8C6: push edx
  loc_0060A8C7: push eax
  loc_0060A8C8: call edi
  loc_0060A8CA: lea ebx, var_28
  loc_0060A8CD: mov eax, var_24
  loc_0060A8D0: push ebx
  loc_0060A8D1: mov edx, 00000008h
  loc_0060A8D6: sub esp, 00000010h
  loc_0060A8D9: mov var_70, edx
  loc_0060A8DC: mov ebx, esp
  loc_0060A8DE: mov ecx, 0043B848h ; "Section4"
  loc_0060A8E3: mov var_68, ecx
  loc_0060A8E6: mov esi, [eax]
  loc_0060A8E8: mov [ebx], edx
  loc_0060A8EA: mov edx, var_6C
  loc_0060A8ED: push eax
  loc_0060A8EE: mov var_AC, eax
  loc_0060A8F4: mov [ebx+00000004h], edx
  loc_0060A8F7: mov [ebx+00000008h], ecx
  loc_0060A8FA: mov ecx, var_64
  loc_0060A8FD: mov [ebx+0000000Ch], ecx
  loc_0060A900: call [esi+0000001Ch]
  loc_0060A903: test eax, eax
  loc_0060A905: fnclex
  loc_0060A907: jge 0060A91Ah
  loc_0060A909: mov edx, var_AC
  loc_0060A90F: push 0000001Ch
  loc_0060A911: push 00439DD8h
  loc_0060A916: push edx
  loc_0060A917: push eax
  loc_0060A918: call edi
  loc_0060A91A: mov eax, var_28
  loc_0060A91D: lea edx, var_2C
  loc_0060A920: push edx
  loc_0060A921: push eax
  loc_0060A922: mov ecx, [eax]
  loc_0060A924: mov esi, eax
  loc_0060A926: call [ecx+0000003Ch]
  loc_0060A929: test eax, eax
  loc_0060A92B: fnclex
  loc_0060A92D: jge 0060A93Ah
  loc_0060A92F: push 0000003Ch
  loc_0060A931: push 00439BF8h
  loc_0060A936: push esi
  loc_0060A937: push eax
  loc_0060A938: call edi
  loc_0060A93A: lea ebx, var_30
  loc_0060A93D: mov eax, var_2C
  loc_0060A940: push ebx
  loc_0060A941: mov edx, 00000008h
  loc_0060A946: sub esp, 00000010h
  loc_0060A949: mov esi, [eax]
  loc_0060A94B: mov ebx, esp
  loc_0060A94D: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060A952: push eax
  loc_0060A953: mov var_BC, eax
  loc_0060A959: mov [ebx], edx
  loc_0060A95B: mov edx, var_7C
  loc_0060A95E: mov [ebx+00000004h], edx
  loc_0060A961: mov [ebx+00000008h], ecx
  loc_0060A964: mov ecx, var_74
  loc_0060A967: mov [ebx+0000000Ch], ecx
  loc_0060A96A: call [esi+0000001Ch]
  loc_0060A96D: test eax, eax
  loc_0060A96F: fnclex
  loc_0060A971: jge 0060A984h
  loc_0060A973: mov edx, var_BC
  loc_0060A979: push 0000001Ch
  loc_0060A97B: push 0043B7D8h
  loc_0060A980: push edx
  loc_0060A981: push eax
  loc_0060A982: call edi
  loc_0060A984: mov ecx, [006680F8h]
  loc_0060A98A: mov edx, [006680FCh]
  loc_0060A990: sub esp, 00000010h
  loc_0060A993: mov eax, esp
  loc_0060A995: push 0043B85Ch ; "Caption"
  loc_0060A99A: mov [eax], ecx
  loc_0060A99C: mov ecx, [00668100h]
  loc_0060A9A2: mov [eax+00000004h], edx
  loc_0060A9A5: mov edx, [00668104h]
  loc_0060A9AB: mov [eax+00000008h], ecx
  loc_0060A9AE: mov [eax+0000000Ch], edx
  loc_0060A9B1: mov eax, var_30
  loc_0060A9B4: push eax
  loc_0060A9B5: call [00401094h] ; __vbaLateMemSt
  loc_0060A9BB: lea ecx, var_30
  loc_0060A9BE: lea edx, var_2C
  loc_0060A9C1: push ecx
  loc_0060A9C2: lea eax, var_28
  loc_0060A9C5: push edx
  loc_0060A9C6: lea ecx, var_24
  loc_0060A9C9: push eax
  loc_0060A9CA: push ecx
  loc_0060A9CB: push 00000004h
  loc_0060A9CD: call [0040104Ch] ; __vbaFreeObjList
  loc_0060A9D3: mov eax, var_C4
  loc_0060A9D9: add esp, 00000014h
  loc_0060A9DC: lea ecx, var_24
  loc_0060A9DF: mov edx, [eax]
  loc_0060A9E1: push ecx
  loc_0060A9E2: push eax
  loc_0060A9E3: call [edx+0000008Ch]
  loc_0060A9E9: test eax, eax
  loc_0060A9EB: fnclex
  loc_0060A9ED: jge 0060AA03h
  loc_0060A9EF: mov edx, var_C4
  loc_0060A9F5: push 0000008Ch
  loc_0060A9FA: push 0043B5D0h
  loc_0060A9FF: push edx
  loc_0060AA00: push eax
  loc_0060AA01: call edi
  loc_0060AA03: lea ebx, var_28
  loc_0060AA06: mov eax, var_24
  loc_0060AA09: push ebx
  loc_0060AA0A: mov edx, 00000008h
  loc_0060AA0F: sub esp, 00000010h
  loc_0060AA12: mov var_70, edx
  loc_0060AA15: mov ebx, esp
  loc_0060AA17: mov ecx, 0043B848h ; "Section4"
  loc_0060AA1C: mov var_68, ecx
  loc_0060AA1F: mov esi, [eax]
  loc_0060AA21: mov [ebx], edx
  loc_0060AA23: mov edx, var_6C
  loc_0060AA26: push eax
  loc_0060AA27: mov var_AC, eax
  loc_0060AA2D: mov [ebx+00000004h], edx
  loc_0060AA30: mov [ebx+00000008h], ecx
  loc_0060AA33: mov ecx, var_64
  loc_0060AA36: mov [ebx+0000000Ch], ecx
  loc_0060AA39: call [esi+0000001Ch]
  loc_0060AA3C: test eax, eax
  loc_0060AA3E: fnclex
  loc_0060AA40: jge 0060AA53h
  loc_0060AA42: mov edx, var_AC
  loc_0060AA48: push 0000001Ch
  loc_0060AA4A: push 00439DD8h
  loc_0060AA4F: push edx
  loc_0060AA50: push eax
  loc_0060AA51: call edi
  loc_0060AA53: mov eax, var_28
  loc_0060AA56: lea edx, var_2C
  loc_0060AA59: push edx
  loc_0060AA5A: push eax
  loc_0060AA5B: mov ecx, [eax]
  loc_0060AA5D: mov esi, eax
  loc_0060AA5F: call [ecx+0000003Ch]
  loc_0060AA62: test eax, eax
  loc_0060AA64: fnclex
  loc_0060AA66: jge 0060AA73h
  loc_0060AA68: push 0000003Ch
  loc_0060AA6A: push 00439BF8h
  loc_0060AA6F: push esi
  loc_0060AA70: push eax
  loc_0060AA71: call edi
  loc_0060AA73: lea ebx, var_30
  loc_0060AA76: mov eax, var_2C
  loc_0060AA79: push ebx
  loc_0060AA7A: mov edx, 00000008h
  loc_0060AA7F: sub esp, 00000010h
  loc_0060AA82: mov esi, [eax]
  loc_0060AA84: mov ebx, esp
  loc_0060AA86: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060AA8B: push eax
  loc_0060AA8C: mov var_BC, eax
  loc_0060AA92: mov [ebx], edx
  loc_0060AA94: mov edx, var_7C
  loc_0060AA97: mov [ebx+00000004h], edx
  loc_0060AA9A: mov [ebx+00000008h], ecx
  loc_0060AA9D: mov ecx, var_74
  loc_0060AAA0: mov [ebx+0000000Ch], ecx
  loc_0060AAA3: call [esi+0000001Ch]
  loc_0060AAA6: test eax, eax
  loc_0060AAA8: fnclex
  loc_0060AAAA: jge 0060AABDh
  loc_0060AAAC: mov edx, var_BC
  loc_0060AAB2: push 0000001Ch
  loc_0060AAB4: push 0043B7D8h
  loc_0060AAB9: push edx
  loc_0060AABA: push eax
  loc_0060AABB: call edi
  loc_0060AABD: mov ecx, [00668108h]
  loc_0060AAC3: mov edx, [0066810Ch]
  loc_0060AAC9: sub esp, 00000010h
  loc_0060AACC: mov eax, esp
  loc_0060AACE: push 0043B85Ch ; "Caption"
  loc_0060AAD3: mov [eax], ecx
  loc_0060AAD5: mov ecx, [00668110h]
  loc_0060AADB: mov [eax+00000004h], edx
  loc_0060AADE: mov edx, [00668114h]
  loc_0060AAE4: mov [eax+00000008h], ecx
  loc_0060AAE7: mov [eax+0000000Ch], edx
  loc_0060AAEA: mov eax, var_30
  loc_0060AAED: push eax
  loc_0060AAEE: call [00401094h] ; __vbaLateMemSt
  loc_0060AAF4: lea ecx, var_30
  loc_0060AAF7: lea edx, var_2C
  loc_0060AAFA: push ecx
  loc_0060AAFB: lea eax, var_28
  loc_0060AAFE: push edx
  loc_0060AAFF: lea ecx, var_24
  loc_0060AB02: push eax
  loc_0060AB03: push ecx
  loc_0060AB04: push 00000004h
  loc_0060AB06: call [0040104Ch] ; __vbaFreeObjList
  loc_0060AB0C: mov eax, var_C4
  loc_0060AB12: add esp, 00000014h
  loc_0060AB15: mov edx, [eax]
  loc_0060AB17: push 0000000Ah
  loc_0060AB19: push eax
  loc_0060AB1A: call [edx+00000048h]
  loc_0060AB1D: test eax, eax
  loc_0060AB1F: fnclex
  loc_0060AB21: jge 0060AB34h
  loc_0060AB23: mov ecx, var_C4
  loc_0060AB29: push 00000048h
  loc_0060AB2B: push 0043B5D0h
  loc_0060AB30: push ecx
  loc_0060AB31: push eax
  loc_0060AB32: call edi
  loc_0060AB34: mov eax, var_C4
  loc_0060AB3A: push 0000000Ah
  loc_0060AB3C: push eax
  loc_0060AB3D: mov edx, [eax]
  loc_0060AB3F: call [edx+00000058h]
  loc_0060AB42: test eax, eax
  loc_0060AB44: fnclex
  loc_0060AB46: jge 0060AB59h
  loc_0060AB48: mov ecx, var_C4
  loc_0060AB4E: push 00000058h
  loc_0060AB50: push 0043B5D0h
  loc_0060AB55: push ecx
  loc_0060AB56: push eax
  loc_0060AB57: call edi
  loc_0060AB59: sub esp, 00000010h
  loc_0060AB5C: mov eax, 00000002h
  loc_0060AB61: mov edx, esp
  loc_0060AB63: mov ecx, eax
  loc_0060AB65: mov var_70, ecx
  loc_0060AB68: mov var_68, eax
  loc_0060AB6B: mov [edx], ecx
  loc_0060AB6D: mov ecx, var_6C
  loc_0060AB70: push 8001000Eh
  loc_0060AB75: mov [edx+00000004h], ecx
  loc_0060AB78: mov ecx, var_C4
  loc_0060AB7E: push ecx
  loc_0060AB7F: mov [edx+00000008h], eax
  loc_0060AB82: mov eax, var_64
  loc_0060AB85: mov [edx+0000000Ch], eax
  loc_0060AB88: call [00401280h] ; __vbaLateIdSt
  loc_0060AB8E: mov edx, var_C4
  loc_0060AB94: push 00000000h
  loc_0060AB96: push 80011003h
  loc_0060AB9B: push edx
  loc_0060AB9C: call [0040102Ch] ; __vbaLateIdCall
  loc_0060ABA2: add esp, 0000000Ch
  loc_0060ABA5: lea eax, var_C4
  loc_0060ABAB: push 00000000h
  loc_0060ABAD: push eax
  loc_0060ABAE: call [004010C8h] ; __vbaObjSetAddref
  loc_0060ABB4: push 0060AC1Eh
  loc_0060ABB9: jmp 0060ABF4h
  loc_0060ABBB: lea ecx, var_20
  loc_0060ABBE: call [0040129Ch] ; __vbaFreeStr
  loc_0060ABC4: lea ecx, var_30
  loc_0060ABC7: lea edx, var_2C
  loc_0060ABCA: push ecx
  loc_0060ABCB: lea eax, var_28
  loc_0060ABCE: push edx
  loc_0060ABCF: lea ecx, var_24
  loc_0060ABD2: push eax
  loc_0060ABD3: push ecx
  loc_0060ABD4: push 00000004h
  loc_0060ABD6: call [0040104Ch] ; __vbaFreeObjList
  loc_0060ABDC: lea edx, var_60
  loc_0060ABDF: lea eax, var_50
  loc_0060ABE2: push edx
  loc_0060ABE3: lea ecx, var_40
  loc_0060ABE6: push eax
  loc_0060ABE7: push ecx
  loc_0060ABE8: push 00000003h
  loc_0060ABEA: call [0040103Ch] ; __vbaFreeVarList
  loc_0060ABF0: add esp, 00000024h
  loc_0060ABF3: ret
  loc_0060ABF4: lea edx, var_C8
  loc_0060ABFA: lea eax, var_C4
  loc_0060AC00: push edx
  loc_0060AC01: push eax
  loc_0060AC02: push 00000002h
  loc_0060AC04: call [0040104Ch] ; __vbaFreeObjList
  loc_0060AC0A: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060AC10: add esp, 0000000Ch
  loc_0060AC13: lea ecx, var_18
  loc_0060AC16: call __vbaFreeStr
  loc_0060AC18: lea ecx, var_1C
  loc_0060AC1B: call __vbaFreeStr
  loc_0060AC1D: ret
  loc_0060AC1E: mov ecx, var_10
  loc_0060AC21: pop edi
  loc_0060AC22: pop esi
  loc_0060AC23: mov fs:[00000000h], ecx
  loc_0060AC2A: pop ebx
  loc_0060AC2B: mov esp, ebp
  loc_0060AC2D: pop ebp
  loc_0060AC2E: ret
End Sub

Private Sub Proc_34_1_60AC40
  loc_0060AC40: push ebp
  loc_0060AC41: mov ebp, esp
  loc_0060AC43: sub esp, 00000008h
  loc_0060AC46: push 00405696h ; __vbaExceptHandler
  loc_0060AC4B: mov eax, fs:[00000000h]
  loc_0060AC51: push eax
  loc_0060AC52: mov fs:[00000000h], esp
  loc_0060AC59: sub esp, 000000D0h
  loc_0060AC5F: push ebx
  loc_0060AC60: push esi
  loc_0060AC61: push edi
  loc_0060AC62: mov var_8, esp
  loc_0060AC65: mov var_4, 004044B0h
  loc_0060AC6C: xor ebx, ebx
  loc_0060AC6E: mov var_18, ebx
  loc_0060AC71: mov var_1C, ebx
  loc_0060AC74: mov var_20, ebx
  loc_0060AC77: mov var_24, ebx
  loc_0060AC7A: mov var_28, ebx
  loc_0060AC7D: mov var_2C, ebx
  loc_0060AC80: mov var_30, ebx
  loc_0060AC83: mov var_40, ebx
  loc_0060AC86: mov var_50, ebx
  loc_0060AC89: mov var_60, ebx
  loc_0060AC8C: mov var_70, ebx
  loc_0060AC8F: mov var_80, ebx
  loc_0060AC92: mov var_90, ebx
  loc_0060AC98: mov var_A4, ebx
  loc_0060AC9E: mov var_C4, ebx
  loc_0060ACA4: mov var_C8, ebx
  loc_0060ACAA: call 0055B320h
  loc_0060ACAF: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0060ACB5: mov edx, 0042DFECh
  loc_0060ACBA: lea ecx, var_18
  loc_0060ACBD: call __vbaStrCopy
  loc_0060ACBF: mov edx, 0043B8D4h ; " SELECT * FROM pemasok ORDER BY kode_pemasok"
  loc_0060ACC4: lea ecx, var_18
  loc_0060ACC7: call __vbaStrCopy
  loc_0060ACC9: push 0042DF30h
  loc_0060ACCE: call [00401168h] ; __vbaNew
  loc_0060ACD4: push eax
  loc_0060ACD5: push 00668090h
  loc_0060ACDA: call [004010B8h] ; __vbaObjSet
  loc_0060ACE0: cmp [0066803Ch], ebx
  loc_0060ACE6: jnz 0060ACF8h
  loc_0060ACE8: push 0066803Ch
  loc_0060ACED: push 0042DEFCh
  loc_0060ACF2: call [004011E8h] ; __vbaNew2
  loc_0060ACF8: mov esi, [0066803Ch]
  loc_0060ACFE: lea ecx, var_40
  loc_0060AD01: mov var_38, 80020004h
  loc_0060AD08: mov var_40, 0000000Ah
  loc_0060AD0F: call [0040123Ch] ; __vbaFreeVarg
  loc_0060AD15: mov eax, [esi]
  loc_0060AD17: lea ecx, var_24
  loc_0060AD1A: push ecx
  loc_0060AD1B: mov ecx, var_18
  loc_0060AD1E: lea edx, var_40
  loc_0060AD21: push 00000001h
  loc_0060AD23: push edx
  loc_0060AD24: push ecx
  loc_0060AD25: push esi
  loc_0060AD26: call [eax+00000040h]
  loc_0060AD29: cmp eax, ebx
  loc_0060AD2B: fnclex
  loc_0060AD2D: jge 0060AD42h
  loc_0060AD2F: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060AD35: push 00000040h
  loc_0060AD37: push 0042E1B0h
  loc_0060AD3C: push esi
  loc_0060AD3D: push eax
  loc_0060AD3E: call edi
  loc_0060AD40: jmp 0060AD48h
  loc_0060AD42: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060AD48: mov edx, var_24
  loc_0060AD4B: mov esi, [00401268h] ; __vbaCastObj
  loc_0060AD51: push 0042DF20h
  loc_0060AD56: push edx
  loc_0060AD57: call __vbaCastObj
  loc_0060AD59: push eax
  loc_0060AD5A: push 00668090h
  loc_0060AD5F: call [004010B8h] ; __vbaObjSet
  loc_0060AD65: lea ecx, var_24
  loc_0060AD68: call [004012A0h] ; __vbaFreeObj
  loc_0060AD6E: lea ecx, var_40
  loc_0060AD71: call [00401020h] ; __vbaFreeVar
  loc_0060AD77: cmp [0066851Ch], ebx
  loc_0060AD7D: jnz 0060AD8Fh
  loc_0060AD7F: push 0066851Ch
  loc_0060AD84: push 00407518h
  loc_0060AD89: call [004011E8h] ; __vbaNew2
  loc_0060AD8F: mov eax, [0066851Ch]
  loc_0060AD94: lea ecx, var_C4
  loc_0060AD9A: push eax
  loc_0060AD9B: push ecx
  loc_0060AD9C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060ADA2: mov edx, var_C4
  loc_0060ADA8: push 0043B7B8h
  loc_0060ADAD: push ebx
  loc_0060ADAE: mov edx, [edx]
  loc_0060ADB0: mov var_E0, edx
  loc_0060ADB6: call __vbaCastObj
  loc_0060ADB8: push eax
  loc_0060ADB9: lea eax, var_24
  loc_0060ADBC: push eax
  loc_0060ADBD: call [004010B8h] ; __vbaObjSet
  loc_0060ADC3: mov ecx, var_C4
  loc_0060ADC9: mov edx, var_E0
  loc_0060ADCF: push eax
  loc_0060ADD0: push ecx
  loc_0060ADD1: call [edx+00000028h]
  loc_0060ADD4: cmp eax, ebx
  loc_0060ADD6: fnclex
  loc_0060ADD8: jge 0060ADEBh
  loc_0060ADDA: mov ecx, var_C4
  loc_0060ADE0: push 00000028h
  loc_0060ADE2: push 0043B5D0h
  loc_0060ADE7: push ecx
  loc_0060ADE8: push eax
  loc_0060ADE9: call edi
  loc_0060ADEB: lea ecx, var_24
  loc_0060ADEE: call [004012A0h] ; __vbaFreeObj
  loc_0060ADF4: mov eax, var_C4
  loc_0060ADFA: push 0042DFECh
  loc_0060ADFF: push eax
  loc_0060AE00: mov edx, [eax]
  loc_0060AE02: call [edx+00000030h]
  loc_0060AE05: cmp eax, ebx
  loc_0060AE07: fnclex
  loc_0060AE09: jge 0060AE1Ch
  loc_0060AE0B: mov ecx, var_C4
  loc_0060AE11: push 00000030h
  loc_0060AE13: push 0043B5D0h
  loc_0060AE18: push ecx
  loc_0060AE19: push eax
  loc_0060AE1A: call edi
  loc_0060AE1C: mov eax, [00668090h]
  loc_0060AE21: lea ecx, var_24
  loc_0060AE24: push ecx
  loc_0060AE25: push eax
  loc_0060AE26: mov edx, [eax]
  loc_0060AE28: call [edx+00000114h]
  loc_0060AE2E: cmp eax, ebx
  loc_0060AE30: fnclex
  loc_0060AE32: jge 0060AE48h
  loc_0060AE34: mov edx, [00668090h]
  loc_0060AE3A: push 00000114h
  loc_0060AE3F: push 0043B7C8h
  loc_0060AE44: push edx
  loc_0060AE45: push eax
  loc_0060AE46: call edi
  loc_0060AE48: mov eax, var_C4
  loc_0060AE4E: mov ecx, var_24
  loc_0060AE51: push 0043B7B8h
  loc_0060AE56: push ecx
  loc_0060AE57: mov esi, [eax]
  loc_0060AE59: call [00401268h] ; __vbaCastObj
  loc_0060AE5F: lea edx, var_28
  loc_0060AE62: push eax
  loc_0060AE63: push edx
  loc_0060AE64: call [004010B8h] ; __vbaObjSet
  loc_0060AE6A: push eax
  loc_0060AE6B: mov eax, var_C4
  loc_0060AE71: push eax
  loc_0060AE72: call [esi+00000028h]
  loc_0060AE75: cmp eax, ebx
  loc_0060AE77: fnclex
  loc_0060AE79: jge 0060AE8Ch
  loc_0060AE7B: mov ecx, var_C4
  loc_0060AE81: push 00000028h
  loc_0060AE83: push 0043B5D0h
  loc_0060AE88: push ecx
  loc_0060AE89: push eax
  loc_0060AE8A: call edi
  loc_0060AE8C: lea edx, var_28
  loc_0060AE8F: lea eax, var_24
  loc_0060AE92: push edx
  loc_0060AE93: push eax
  loc_0060AE94: push 00000002h
  loc_0060AE96: call [0040104Ch] ; __vbaFreeObjList
  loc_0060AE9C: mov eax, var_C4
  loc_0060AEA2: add esp, 0000000Ch
  loc_0060AEA5: lea edx, var_24
  loc_0060AEA8: mov ecx, [eax]
  loc_0060AEAA: push edx
  loc_0060AEAB: push eax
  loc_0060AEAC: call [ecx+0000008Ch]
  loc_0060AEB2: cmp eax, ebx
  loc_0060AEB4: fnclex
  loc_0060AEB6: jge 0060AECCh
  loc_0060AEB8: mov ecx, var_C4
  loc_0060AEBE: push 0000008Ch
  loc_0060AEC3: push 0043B5D0h
  loc_0060AEC8: push ecx
  loc_0060AEC9: push eax
  loc_0060AECA: call edi
  loc_0060AECC: lea esi, var_28
  loc_0060AECF: mov eax, var_24
  loc_0060AED2: push esi
  loc_0060AED3: mov ecx, 00000008h
  loc_0060AED8: sub esp, 00000010h
  loc_0060AEDB: mov var_70, ecx
  loc_0060AEDE: mov esi, esp
  loc_0060AEE0: mov var_68, 00439FB0h ; "Section1"
  loc_0060AEE7: mov edx, [eax]
  loc_0060AEE9: push eax
  loc_0060AEEA: mov [esi], ecx
  loc_0060AEEC: mov ecx, var_6C
  loc_0060AEEF: mov var_AC, eax
  loc_0060AEF5: mov [esi+00000004h], ecx
  loc_0060AEF8: mov ecx, var_68
  loc_0060AEFB: mov [esi+00000008h], ecx
  loc_0060AEFE: mov ecx, var_64
  loc_0060AF01: mov [esi+0000000Ch], ecx
  loc_0060AF04: call [edx+0000001Ch]
  loc_0060AF07: cmp eax, ebx
  loc_0060AF09: fnclex
  loc_0060AF0B: jge 0060AF1Eh
  loc_0060AF0D: mov edx, var_AC
  loc_0060AF13: push 0000001Ch
  loc_0060AF15: push 00439DD8h
  loc_0060AF1A: push edx
  loc_0060AF1B: push eax
  loc_0060AF1C: call edi
  loc_0060AF1E: mov eax, var_28
  loc_0060AF21: lea edx, var_2C
  loc_0060AF24: push edx
  loc_0060AF25: push eax
  loc_0060AF26: mov ecx, [eax]
  loc_0060AF28: mov esi, eax
  loc_0060AF2A: call [ecx+0000003Ch]
  loc_0060AF2D: cmp eax, ebx
  loc_0060AF2F: fnclex
  loc_0060AF31: jge 0060AF3Eh
  loc_0060AF33: push 0000003Ch
  loc_0060AF35: push 00439BF8h
  loc_0060AF3A: push esi
  loc_0060AF3B: push eax
  loc_0060AF3C: call edi
  loc_0060AF3E: mov eax, var_2C
  loc_0060AF41: mov var_2C, ebx
  loc_0060AF44: push eax
  loc_0060AF45: lea eax, var_C8
  loc_0060AF4B: push eax
  loc_0060AF4C: call [004010B8h] ; __vbaObjSet
  loc_0060AF52: lea ecx, var_28
  loc_0060AF55: lea edx, var_24
  loc_0060AF58: push ecx
  loc_0060AF59: push edx
  loc_0060AF5A: push 00000002h
  loc_0060AF5C: call [0040104Ch] ; __vbaFreeObjList
  loc_0060AF62: mov eax, var_C8
  loc_0060AF68: add esp, 0000000Ch
  loc_0060AF6B: lea edx, var_A4
  loc_0060AF71: mov ecx, [eax]
  loc_0060AF73: push edx
  loc_0060AF74: push eax
  loc_0060AF75: call [ecx+00000020h]
  loc_0060AF78: cmp eax, ebx
  loc_0060AF7A: fnclex
  loc_0060AF7C: jge 0060AF8Fh
  loc_0060AF7E: mov ecx, var_C8
  loc_0060AF84: push 00000020h
  loc_0060AF86: push 0043B7D8h
  loc_0060AF8B: push ecx
  loc_0060AF8C: push eax
  loc_0060AF8D: call edi
  loc_0060AF8F: mov ecx, var_A4
  loc_0060AF95: call [00401138h] ; __vbaI2I4
  loc_0060AF9B: mov var_D0, eax
  loc_0060AFA1: mov ebx, 00000001h
  loc_0060AFA6: cmp bx, var_D0
  loc_0060AFAD: mov var_14, ebx
  loc_0060AFB0: jg 0060B221h
  loc_0060AFB6: lea esi, var_24
  loc_0060AFB9: mov ecx, var_C8
  loc_0060AFBF: push esi
  loc_0060AFC0: mov eax, 00000002h
  loc_0060AFC5: sub esp, 00000010h
  loc_0060AFC8: mov var_70, eax
  loc_0060AFCB: mov esi, esp
  loc_0060AFCD: mov var_68, bx
  loc_0060AFD1: mov edx, [ecx]
  loc_0060AFD3: push ecx
  loc_0060AFD4: mov [esi], eax
  loc_0060AFD6: mov eax, var_6C
  loc_0060AFD9: mov [esi+00000004h], eax
  loc_0060AFDC: mov eax, var_68
  loc_0060AFDF: mov [esi+00000008h], eax
  loc_0060AFE2: mov eax, var_64
  loc_0060AFE5: mov [esi+0000000Ch], eax
  loc_0060AFE8: call [edx+0000001Ch]
  loc_0060AFEB: test eax, eax
  loc_0060AFED: fnclex
  loc_0060AFEF: jge 0060B002h
  loc_0060AFF1: mov ecx, var_C8
  loc_0060AFF7: push 0000001Ch
  loc_0060AFF9: push 0043B7D8h
  loc_0060AFFE: push ecx
  loc_0060AFFF: push eax
  loc_0060B000: call edi
  loc_0060B002: mov edx, var_24
  loc_0060B005: push 0043B7E8h
  loc_0060B00A: push edx
  loc_0060B00B: call [004011CCh] ; __vbaCheckType
  loc_0060B011: lea ecx, var_24
  loc_0060B014: mov si, ax
  loc_0060B017: call [004012A0h] ; __vbaFreeObj
  loc_0060B01D: test si, si
  loc_0060B020: Unknown_795()
  loc_0060B026: mov var_68, bx
  loc_0060B02A: lea ebx, var_24
  loc_0060B02D: push ebx
  loc_0060B02E: mov ecx, var_C8
  loc_0060B034: sub esp, 00000010h
  loc_0060B037: mov eax, 00000002h
  loc_0060B03C: mov ebx, esp
  loc_0060B03E: mov var_70, eax
  loc_0060B041: mov edx, [ecx]
  loc_0060B043: mov esi, 0042DFECh
  loc_0060B048: mov [ebx], eax
  loc_0060B04A: mov eax, var_6C
  loc_0060B04D: push ecx
  loc_0060B04E: mov var_78, esi
  loc_0060B051: mov [ebx+00000004h], eax
  loc_0060B054: mov eax, var_68
  loc_0060B057: mov edi, 00000008h
  loc_0060B05C: mov [ebx+00000008h], eax
  loc_0060B05F: mov eax, var_64
  loc_0060B062: mov [ebx+0000000Ch], eax
  loc_0060B065: call [edx+0000001Ch]
  loc_0060B068: test eax, eax
  loc_0060B06A: fnclex
  loc_0060B06C: jge 0060B083h
  loc_0060B06E: mov ecx, var_C8
  loc_0060B074: push 0000001Ch
  loc_0060B076: push 0043B7D8h
  loc_0060B07B: push ecx
  loc_0060B07C: push eax
  loc_0060B07D: call [00401080h] ; __vbaHresultCheckObj
  loc_0060B083: mov eax, var_7C
  loc_0060B086: sub esp, 00000010h
  loc_0060B089: mov edx, esp
  loc_0060B08B: mov ecx, var_74
  loc_0060B08E: push 0043B7F8h ; "DataMember"
  loc_0060B093: mov [edx], edi
  loc_0060B095: mov [edx+00000004h], eax
  loc_0060B098: mov [edx+00000008h], esi
  loc_0060B09B: mov [edx+0000000Ch], ecx
  loc_0060B09E: mov edx, var_24
  loc_0060B0A1: push edx
  loc_0060B0A2: call [00401094h] ; __vbaLateMemSt
  loc_0060B0A8: lea ecx, var_24
  loc_0060B0AB: call [004012A0h] ; __vbaFreeObj
  loc_0060B0B1: mov eax, [00668090h]
  loc_0060B0B6: lea edx, var_24
  loc_0060B0B9: push edx
  loc_0060B0BA: push eax
  loc_0060B0BB: mov ecx, [eax]
  loc_0060B0BD: call [ecx+00000054h]
  loc_0060B0C0: test eax, eax
  loc_0060B0C2: fnclex
  loc_0060B0C4: jge 0060B0DBh
  loc_0060B0C6: mov ecx, [00668090h]
  loc_0060B0CC: push 00000054h
  loc_0060B0CE: push 0042DF44h
  loc_0060B0D3: push ecx
  loc_0060B0D4: push eax
  loc_0060B0D5: call [00401080h] ; __vbaHresultCheckObj
  loc_0060B0DB: mov ebx, var_14
  loc_0060B0DE: lea edi, var_28
  loc_0060B0E1: mov dx, bx
  loc_0060B0E4: push edi
  loc_0060B0E5: sub dx, 0001h
  loc_0060B0E9: mov eax, var_24
  loc_0060B0EC: jo 0060B9AFh
  loc_0060B0F2: sub esp, 00000010h
  loc_0060B0F5: mov ecx, 00000002h
  loc_0060B0FA: mov edi, esp
  loc_0060B0FC: mov var_70, ecx
  loc_0060B0FF: mov var_68, dx
  loc_0060B103: mov edx, [eax]
  loc_0060B105: mov [edi], ecx
  loc_0060B107: mov ecx, var_6C
  loc_0060B10A: push eax
  loc_0060B10B: mov esi, eax
  loc_0060B10D: mov [edi+00000004h], ecx
  loc_0060B110: mov ecx, var_68
  loc_0060B113: mov [edi+00000008h], ecx
  loc_0060B116: mov ecx, var_64
  loc_0060B119: mov [edi+0000000Ch], ecx
  loc_0060B11C: call [edx+00000028h]
  loc_0060B11F: test eax, eax
  loc_0060B121: fnclex
  loc_0060B123: jge 0060B138h
  loc_0060B125: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060B12B: push 00000028h
  loc_0060B12D: push 0042DFACh
  loc_0060B132: push esi
  loc_0060B133: push eax
  loc_0060B134: call edi
  loc_0060B136: jmp 0060B13Eh
  loc_0060B138: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060B13E: mov eax, var_28
  loc_0060B141: lea ecx, var_20
  loc_0060B144: push ecx
  loc_0060B145: push eax
  loc_0060B146: mov edx, [eax]
  loc_0060B148: mov esi, eax
  loc_0060B14A: call [edx+0000002Ch]
  loc_0060B14D: test eax, eax
  loc_0060B14F: fnclex
  loc_0060B151: jge 0060B15Eh
  loc_0060B153: push 0000002Ch
  loc_0060B155: push 0042DFBCh
  loc_0060B15A: push esi
  loc_0060B15B: push eax
  loc_0060B15C: call edi
  loc_0060B15E: mov eax, var_20
  loc_0060B161: lea esi, var_2C
  loc_0060B164: push esi
  loc_0060B165: mov ecx, var_C8
  loc_0060B16B: sub esp, 00000010h
  loc_0060B16E: mov var_38, eax
  loc_0060B171: mov esi, esp
  loc_0060B173: mov eax, 00000002h
  loc_0060B178: mov var_78, bx
  loc_0060B17C: mov var_20, 00000000h
  loc_0060B183: mov [esi], eax
  loc_0060B185: mov eax, var_7C
  loc_0060B188: mov var_40, 00000008h
  loc_0060B18F: mov edx, [ecx]
  loc_0060B191: mov [esi+00000004h], eax
  loc_0060B194: mov eax, var_78
  loc_0060B197: push ecx
  loc_0060B198: mov [esi+00000008h], eax
  loc_0060B19B: mov eax, var_74
  loc_0060B19E: mov [esi+0000000Ch], eax
  loc_0060B1A1: call [edx+0000001Ch]
  loc_0060B1A4: test eax, eax
  loc_0060B1A6: fnclex
  loc_0060B1A8: jge 0060B1BBh
  loc_0060B1AA: mov ecx, var_C8
  loc_0060B1B0: push 0000001Ch
  loc_0060B1B2: push 0043B7D8h
  loc_0060B1B7: push ecx
  loc_0060B1B8: push eax
  loc_0060B1B9: call edi
  loc_0060B1BB: mov eax, var_40
  loc_0060B1BE: mov ecx, var_3C
  loc_0060B1C1: sub esp, 00000010h
  loc_0060B1C4: mov edx, esp
  loc_0060B1C6: push 0043B810h ; "DataField"
  loc_0060B1CB: mov [edx], eax
  loc_0060B1CD: mov eax, var_38
  loc_0060B1D0: mov [edx+00000004h], ecx
  loc_0060B1D3: mov ecx, var_34
  loc_0060B1D6: mov [edx+00000008h], eax
  loc_0060B1D9: mov [edx+0000000Ch], ecx
  loc_0060B1DC: mov edx, var_2C
  loc_0060B1DF: push edx
  loc_0060B1E0: call [00401094h] ; __vbaLateMemSt
  loc_0060B1E6: lea eax, var_2C
  loc_0060B1E9: lea ecx, var_28
  loc_0060B1EC: push eax
  loc_0060B1ED: lea edx, var_24
  loc_0060B1F0: push ecx
  loc_0060B1F1: push edx
  loc_0060B1F2: push 00000003h
  loc_0060B1F4: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B1FA: add esp, 00000010h
  loc_0060B1FD: lea ecx, var_40
  loc_0060B200: call [00401020h] ; __vbaFreeVar
  loc_0060B206: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060B20C: mov eax, 00000001h
  loc_0060B211: add ax, bx
  loc_0060B214: jo 0060B9AFh
  loc_0060B21A: mov ebx, eax
  loc_0060B21C: jmp 0060AFA6h
  loc_0060B221: lea eax, var_C8
  loc_0060B227: push 00000000h
  loc_0060B229: push eax
  loc_0060B22A: call [004010C8h] ; __vbaObjSetAddref
  loc_0060B230: lea ecx, var_40
  loc_0060B233: push ecx
  loc_0060B234: call [00401220h] ; rtcGetDateVar
  loc_0060B23A: lea edx, var_70
  loc_0060B23D: lea ecx, var_50
  loc_0060B240: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060B247: mov var_70, 00000008h
  loc_0060B24E: call [00401238h] ; __vbaVarDup
  loc_0060B254: push 00000001h
  loc_0060B256: lea edx, var_50
  loc_0060B259: push 00000001h
  loc_0060B25B: lea eax, var_40
  loc_0060B25E: push edx
  loc_0060B25F: lea ecx, var_60
  loc_0060B262: push eax
  loc_0060B263: push ecx
  loc_0060B264: call [00401078h] ; rtcVarFromFormatVar
  loc_0060B26A: mov eax, var_C4
  loc_0060B270: lea ecx, var_24
  loc_0060B273: push ecx
  loc_0060B274: push eax
  loc_0060B275: mov edx, [eax]
  loc_0060B277: call [edx+0000008Ch]
  loc_0060B27D: test eax, eax
  loc_0060B27F: fnclex
  loc_0060B281: jge 0060B297h
  loc_0060B283: mov edx, var_C4
  loc_0060B289: push 0000008Ch
  loc_0060B28E: push 0043B5D0h
  loc_0060B293: push edx
  loc_0060B294: push eax
  loc_0060B295: call edi
  loc_0060B297: lea ebx, var_28
  loc_0060B29A: mov eax, var_24
  loc_0060B29D: push ebx
  loc_0060B29E: mov edx, 00000008h
  loc_0060B2A3: sub esp, 00000010h
  loc_0060B2A6: mov esi, [eax]
  loc_0060B2A8: mov ebx, esp
  loc_0060B2AA: mov ecx, 0043B848h ; "Section4"
  loc_0060B2AF: push eax
  loc_0060B2B0: mov var_AC, eax
  loc_0060B2B6: mov [ebx], edx
  loc_0060B2B8: mov edx, var_7C
  loc_0060B2BB: mov [ebx+00000004h], edx
  loc_0060B2BE: mov [ebx+00000008h], ecx
  loc_0060B2C1: mov ecx, var_74
  loc_0060B2C4: mov [ebx+0000000Ch], ecx
  loc_0060B2C7: call [esi+0000001Ch]
  loc_0060B2CA: test eax, eax
  loc_0060B2CC: fnclex
  loc_0060B2CE: jge 0060B2E1h
  loc_0060B2D0: mov edx, var_AC
  loc_0060B2D6: push 0000001Ch
  loc_0060B2D8: push 00439DD8h
  loc_0060B2DD: push edx
  loc_0060B2DE: push eax
  loc_0060B2DF: call edi
  loc_0060B2E1: mov eax, var_28
  loc_0060B2E4: lea edx, var_2C
  loc_0060B2E7: push edx
  loc_0060B2E8: push eax
  loc_0060B2E9: mov ecx, [eax]
  loc_0060B2EB: mov esi, eax
  loc_0060B2ED: call [ecx+0000003Ch]
  loc_0060B2F0: test eax, eax
  loc_0060B2F2: fnclex
  loc_0060B2F4: jge 0060B301h
  loc_0060B2F6: push 0000003Ch
  loc_0060B2F8: push 00439BF8h
  loc_0060B2FD: push esi
  loc_0060B2FE: push eax
  loc_0060B2FF: call edi
  loc_0060B301: lea ebx, var_30
  loc_0060B304: mov eax, var_2C
  loc_0060B307: push ebx
  loc_0060B308: mov edx, 00000008h
  loc_0060B30D: sub esp, 00000010h
  loc_0060B310: mov esi, [eax]
  loc_0060B312: mov ebx, esp
  loc_0060B314: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060B319: push eax
  loc_0060B31A: mov var_BC, eax
  loc_0060B320: mov [ebx], edx
  loc_0060B322: mov edx, var_8C
  loc_0060B328: mov [ebx+00000004h], edx
  loc_0060B32B: mov [ebx+00000008h], ecx
  loc_0060B32E: mov ecx, var_84
  loc_0060B334: mov [ebx+0000000Ch], ecx
  loc_0060B337: call [esi+0000001Ch]
  loc_0060B33A: test eax, eax
  loc_0060B33C: fnclex
  loc_0060B33E: jge 0060B351h
  loc_0060B340: mov edx, var_BC
  loc_0060B346: push 0000001Ch
  loc_0060B348: push 0043B7D8h
  loc_0060B34D: push edx
  loc_0060B34E: push eax
  loc_0060B34F: call edi
  loc_0060B351: mov ecx, var_60
  loc_0060B354: mov edx, var_5C
  loc_0060B357: sub esp, 00000010h
  loc_0060B35A: mov eax, esp
  loc_0060B35C: push 0043B85Ch ; "Caption"
  loc_0060B361: mov [eax], ecx
  loc_0060B363: mov ecx, var_58
  loc_0060B366: mov [eax+00000004h], edx
  loc_0060B369: mov edx, var_54
  loc_0060B36C: mov [eax+00000008h], ecx
  loc_0060B36F: mov [eax+0000000Ch], edx
  loc_0060B372: mov eax, var_30
  loc_0060B375: push eax
  loc_0060B376: call [00401094h] ; __vbaLateMemSt
  loc_0060B37C: lea ecx, var_30
  loc_0060B37F: lea edx, var_2C
  loc_0060B382: push ecx
  loc_0060B383: lea eax, var_28
  loc_0060B386: push edx
  loc_0060B387: lea ecx, var_24
  loc_0060B38A: push eax
  loc_0060B38B: push ecx
  loc_0060B38C: push 00000004h
  loc_0060B38E: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B394: lea edx, var_60
  loc_0060B397: lea eax, var_50
  loc_0060B39A: push edx
  loc_0060B39B: lea ecx, var_40
  loc_0060B39E: push eax
  loc_0060B39F: push ecx
  loc_0060B3A0: push 00000003h
  loc_0060B3A2: call [0040103Ch] ; __vbaFreeVarList
  loc_0060B3A8: mov eax, var_C4
  loc_0060B3AE: add esp, 00000024h
  loc_0060B3B1: lea ecx, var_24
  loc_0060B3B4: mov edx, [eax]
  loc_0060B3B6: push ecx
  loc_0060B3B7: push eax
  loc_0060B3B8: call [edx+0000008Ch]
  loc_0060B3BE: test eax, eax
  loc_0060B3C0: fnclex
  loc_0060B3C2: jge 0060B3D8h
  loc_0060B3C4: mov edx, var_C4
  loc_0060B3CA: push 0000008Ch
  loc_0060B3CF: push 0043B5D0h
  loc_0060B3D4: push edx
  loc_0060B3D5: push eax
  loc_0060B3D6: call edi
  loc_0060B3D8: lea ebx, var_28
  loc_0060B3DB: mov eax, var_24
  loc_0060B3DE: push ebx
  loc_0060B3DF: mov edx, 00000008h
  loc_0060B3E4: sub esp, 00000010h
  loc_0060B3E7: mov var_70, edx
  loc_0060B3EA: mov ebx, esp
  loc_0060B3EC: mov ecx, 0043B848h ; "Section4"
  loc_0060B3F1: mov var_68, ecx
  loc_0060B3F4: mov esi, [eax]
  loc_0060B3F6: mov [ebx], edx
  loc_0060B3F8: mov edx, var_6C
  loc_0060B3FB: push eax
  loc_0060B3FC: mov var_AC, eax
  loc_0060B402: mov [ebx+00000004h], edx
  loc_0060B405: mov [ebx+00000008h], ecx
  loc_0060B408: mov ecx, var_64
  loc_0060B40B: mov [ebx+0000000Ch], ecx
  loc_0060B40E: call [esi+0000001Ch]
  loc_0060B411: test eax, eax
  loc_0060B413: fnclex
  loc_0060B415: jge 0060B428h
  loc_0060B417: mov edx, var_AC
  loc_0060B41D: push 0000001Ch
  loc_0060B41F: push 00439DD8h
  loc_0060B424: push edx
  loc_0060B425: push eax
  loc_0060B426: call edi
  loc_0060B428: mov eax, var_28
  loc_0060B42B: lea edx, var_2C
  loc_0060B42E: push edx
  loc_0060B42F: push eax
  loc_0060B430: mov ecx, [eax]
  loc_0060B432: mov esi, eax
  loc_0060B434: call [ecx+0000003Ch]
  loc_0060B437: test eax, eax
  loc_0060B439: fnclex
  loc_0060B43B: jge 0060B448h
  loc_0060B43D: push 0000003Ch
  loc_0060B43F: push 00439BF8h
  loc_0060B444: push esi
  loc_0060B445: push eax
  loc_0060B446: call edi
  loc_0060B448: lea ebx, var_30
  loc_0060B44B: mov eax, var_2C
  loc_0060B44E: push ebx
  loc_0060B44F: mov edx, 00000008h
  loc_0060B454: sub esp, 00000010h
  loc_0060B457: mov esi, [eax]
  loc_0060B459: mov ebx, esp
  loc_0060B45B: mov ecx, 0043B870h ; "LabelNama"
  loc_0060B460: push eax
  loc_0060B461: mov var_BC, eax
  loc_0060B467: mov [ebx], edx
  loc_0060B469: mov edx, var_7C
  loc_0060B46C: mov [ebx+00000004h], edx
  loc_0060B46F: mov [ebx+00000008h], ecx
  loc_0060B472: mov ecx, var_74
  loc_0060B475: mov [ebx+0000000Ch], ecx
  loc_0060B478: call [esi+0000001Ch]
  loc_0060B47B: test eax, eax
  loc_0060B47D: fnclex
  loc_0060B47F: jge 0060B492h
  loc_0060B481: mov edx, var_BC
  loc_0060B487: push 0000001Ch
  loc_0060B489: push 0043B7D8h
  loc_0060B48E: push edx
  loc_0060B48F: push eax
  loc_0060B490: call edi
  loc_0060B492: mov ecx, [006680D8h]
  loc_0060B498: mov edx, [006680DCh]
  loc_0060B49E: sub esp, 00000010h
  loc_0060B4A1: mov eax, esp
  loc_0060B4A3: push 0043B85Ch ; "Caption"
  loc_0060B4A8: mov [eax], ecx
  loc_0060B4AA: mov ecx, [006680E0h]
  loc_0060B4B0: mov [eax+00000004h], edx
  loc_0060B4B3: mov edx, [006680E4h]
  loc_0060B4B9: mov [eax+00000008h], ecx
  loc_0060B4BC: mov [eax+0000000Ch], edx
  loc_0060B4BF: mov eax, var_30
  loc_0060B4C2: push eax
  loc_0060B4C3: call [00401094h] ; __vbaLateMemSt
  loc_0060B4C9: lea ecx, var_30
  loc_0060B4CC: lea edx, var_2C
  loc_0060B4CF: push ecx
  loc_0060B4D0: lea eax, var_28
  loc_0060B4D3: push edx
  loc_0060B4D4: lea ecx, var_24
  loc_0060B4D7: push eax
  loc_0060B4D8: push ecx
  loc_0060B4D9: push 00000004h
  loc_0060B4DB: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B4E1: mov eax, var_C4
  loc_0060B4E7: add esp, 00000014h
  loc_0060B4EA: lea ecx, var_24
  loc_0060B4ED: mov edx, [eax]
  loc_0060B4EF: push ecx
  loc_0060B4F0: push eax
  loc_0060B4F1: call [edx+0000008Ch]
  loc_0060B4F7: test eax, eax
  loc_0060B4F9: fnclex
  loc_0060B4FB: jge 0060B511h
  loc_0060B4FD: mov edx, var_C4
  loc_0060B503: push 0000008Ch
  loc_0060B508: push 0043B5D0h
  loc_0060B50D: push edx
  loc_0060B50E: push eax
  loc_0060B50F: call edi
  loc_0060B511: lea ebx, var_28
  loc_0060B514: mov eax, var_24
  loc_0060B517: push ebx
  loc_0060B518: mov edx, 00000008h
  loc_0060B51D: sub esp, 00000010h
  loc_0060B520: mov var_70, edx
  loc_0060B523: mov ebx, esp
  loc_0060B525: mov ecx, 0043B848h ; "Section4"
  loc_0060B52A: mov var_68, ecx
  loc_0060B52D: mov esi, [eax]
  loc_0060B52F: mov [ebx], edx
  loc_0060B531: mov edx, var_6C
  loc_0060B534: push eax
  loc_0060B535: mov var_AC, eax
  loc_0060B53B: mov [ebx+00000004h], edx
  loc_0060B53E: mov [ebx+00000008h], ecx
  loc_0060B541: mov ecx, var_64
  loc_0060B544: mov [ebx+0000000Ch], ecx
  loc_0060B547: call [esi+0000001Ch]
  loc_0060B54A: test eax, eax
  loc_0060B54C: fnclex
  loc_0060B54E: jge 0060B561h
  loc_0060B550: mov edx, var_AC
  loc_0060B556: push 0000001Ch
  loc_0060B558: push 00439DD8h
  loc_0060B55D: push edx
  loc_0060B55E: push eax
  loc_0060B55F: call edi
  loc_0060B561: mov eax, var_28
  loc_0060B564: lea edx, var_2C
  loc_0060B567: push edx
  loc_0060B568: push eax
  loc_0060B569: mov ecx, [eax]
  loc_0060B56B: mov esi, eax
  loc_0060B56D: call [ecx+0000003Ch]
  loc_0060B570: test eax, eax
  loc_0060B572: fnclex
  loc_0060B574: jge 0060B581h
  loc_0060B576: push 0000003Ch
  loc_0060B578: push 00439BF8h
  loc_0060B57D: push esi
  loc_0060B57E: push eax
  loc_0060B57F: call edi
  loc_0060B581: lea ebx, var_30
  loc_0060B584: mov eax, var_2C
  loc_0060B587: push ebx
  loc_0060B588: mov edx, 00000008h
  loc_0060B58D: sub esp, 00000010h
  loc_0060B590: mov esi, [eax]
  loc_0060B592: mov ebx, esp
  loc_0060B594: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060B599: push eax
  loc_0060B59A: mov var_BC, eax
  loc_0060B5A0: mov [ebx], edx
  loc_0060B5A2: mov edx, var_7C
  loc_0060B5A5: mov [ebx+00000004h], edx
  loc_0060B5A8: mov [ebx+00000008h], ecx
  loc_0060B5AB: mov ecx, var_74
  loc_0060B5AE: mov [ebx+0000000Ch], ecx
  loc_0060B5B1: call [esi+0000001Ch]
  loc_0060B5B4: test eax, eax
  loc_0060B5B6: fnclex
  loc_0060B5B8: jge 0060B5CBh
  loc_0060B5BA: mov edx, var_BC
  loc_0060B5C0: push 0000001Ch
  loc_0060B5C2: push 0043B7D8h
  loc_0060B5C7: push edx
  loc_0060B5C8: push eax
  loc_0060B5C9: call edi
  loc_0060B5CB: mov ecx, [006680E8h]
  loc_0060B5D1: mov edx, [006680ECh]
  loc_0060B5D7: sub esp, 00000010h
  loc_0060B5DA: mov eax, esp
  loc_0060B5DC: push 0043B85Ch ; "Caption"
  loc_0060B5E1: mov [eax], ecx
  loc_0060B5E3: mov ecx, [006680F0h]
  loc_0060B5E9: mov [eax+00000004h], edx
  loc_0060B5EC: mov edx, [006680F4h]
  loc_0060B5F2: mov [eax+00000008h], ecx
  loc_0060B5F5: mov [eax+0000000Ch], edx
  loc_0060B5F8: mov eax, var_30
  loc_0060B5FB: push eax
  loc_0060B5FC: call [00401094h] ; __vbaLateMemSt
  loc_0060B602: lea ecx, var_30
  loc_0060B605: lea edx, var_2C
  loc_0060B608: push ecx
  loc_0060B609: lea eax, var_28
  loc_0060B60C: push edx
  loc_0060B60D: lea ecx, var_24
  loc_0060B610: push eax
  loc_0060B611: push ecx
  loc_0060B612: push 00000004h
  loc_0060B614: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B61A: mov eax, var_C4
  loc_0060B620: add esp, 00000014h
  loc_0060B623: lea ecx, var_24
  loc_0060B626: mov edx, [eax]
  loc_0060B628: push ecx
  loc_0060B629: push eax
  loc_0060B62A: call [edx+0000008Ch]
  loc_0060B630: test eax, eax
  loc_0060B632: fnclex
  loc_0060B634: jge 0060B64Ah
  loc_0060B636: mov edx, var_C4
  loc_0060B63C: push 0000008Ch
  loc_0060B641: push 0043B5D0h
  loc_0060B646: push edx
  loc_0060B647: push eax
  loc_0060B648: call edi
  loc_0060B64A: lea ebx, var_28
  loc_0060B64D: mov eax, var_24
  loc_0060B650: push ebx
  loc_0060B651: mov edx, 00000008h
  loc_0060B656: sub esp, 00000010h
  loc_0060B659: mov var_70, edx
  loc_0060B65C: mov ebx, esp
  loc_0060B65E: mov ecx, 0043B848h ; "Section4"
  loc_0060B663: mov var_68, ecx
  loc_0060B666: mov esi, [eax]
  loc_0060B668: mov [ebx], edx
  loc_0060B66A: mov edx, var_6C
  loc_0060B66D: push eax
  loc_0060B66E: mov var_AC, eax
  loc_0060B674: mov [ebx+00000004h], edx
  loc_0060B677: mov [ebx+00000008h], ecx
  loc_0060B67A: mov ecx, var_64
  loc_0060B67D: mov [ebx+0000000Ch], ecx
  loc_0060B680: call [esi+0000001Ch]
  loc_0060B683: test eax, eax
  loc_0060B685: fnclex
  loc_0060B687: jge 0060B69Ah
  loc_0060B689: mov edx, var_AC
  loc_0060B68F: push 0000001Ch
  loc_0060B691: push 00439DD8h
  loc_0060B696: push edx
  loc_0060B697: push eax
  loc_0060B698: call edi
  loc_0060B69A: mov eax, var_28
  loc_0060B69D: lea edx, var_2C
  loc_0060B6A0: push edx
  loc_0060B6A1: push eax
  loc_0060B6A2: mov ecx, [eax]
  loc_0060B6A4: mov esi, eax
  loc_0060B6A6: call [ecx+0000003Ch]
  loc_0060B6A9: test eax, eax
  loc_0060B6AB: fnclex
  loc_0060B6AD: jge 0060B6BAh
  loc_0060B6AF: push 0000003Ch
  loc_0060B6B1: push 00439BF8h
  loc_0060B6B6: push esi
  loc_0060B6B7: push eax
  loc_0060B6B8: call edi
  loc_0060B6BA: lea ebx, var_30
  loc_0060B6BD: mov eax, var_2C
  loc_0060B6C0: push ebx
  loc_0060B6C1: mov edx, 00000008h
  loc_0060B6C6: sub esp, 00000010h
  loc_0060B6C9: mov esi, [eax]
  loc_0060B6CB: mov ebx, esp
  loc_0060B6CD: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060B6D2: push eax
  loc_0060B6D3: mov var_BC, eax
  loc_0060B6D9: mov [ebx], edx
  loc_0060B6DB: mov edx, var_7C
  loc_0060B6DE: mov [ebx+00000004h], edx
  loc_0060B6E1: mov [ebx+00000008h], ecx
  loc_0060B6E4: mov ecx, var_74
  loc_0060B6E7: mov [ebx+0000000Ch], ecx
  loc_0060B6EA: call [esi+0000001Ch]
  loc_0060B6ED: test eax, eax
  loc_0060B6EF: fnclex
  loc_0060B6F1: jge 0060B704h
  loc_0060B6F3: mov edx, var_BC
  loc_0060B6F9: push 0000001Ch
  loc_0060B6FB: push 0043B7D8h
  loc_0060B700: push edx
  loc_0060B701: push eax
  loc_0060B702: call edi
  loc_0060B704: mov ecx, [006680F8h]
  loc_0060B70A: mov edx, [006680FCh]
  loc_0060B710: sub esp, 00000010h
  loc_0060B713: mov eax, esp
  loc_0060B715: push 0043B85Ch ; "Caption"
  loc_0060B71A: mov [eax], ecx
  loc_0060B71C: mov ecx, [00668100h]
  loc_0060B722: mov [eax+00000004h], edx
  loc_0060B725: mov edx, [00668104h]
  loc_0060B72B: mov [eax+00000008h], ecx
  loc_0060B72E: mov [eax+0000000Ch], edx
  loc_0060B731: mov eax, var_30
  loc_0060B734: push eax
  loc_0060B735: call [00401094h] ; __vbaLateMemSt
  loc_0060B73B: lea ecx, var_30
  loc_0060B73E: lea edx, var_2C
  loc_0060B741: push ecx
  loc_0060B742: lea eax, var_28
  loc_0060B745: push edx
  loc_0060B746: lea ecx, var_24
  loc_0060B749: push eax
  loc_0060B74A: push ecx
  loc_0060B74B: push 00000004h
  loc_0060B74D: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B753: mov eax, var_C4
  loc_0060B759: add esp, 00000014h
  loc_0060B75C: lea ecx, var_24
  loc_0060B75F: mov edx, [eax]
  loc_0060B761: push ecx
  loc_0060B762: push eax
  loc_0060B763: call [edx+0000008Ch]
  loc_0060B769: test eax, eax
  loc_0060B76B: fnclex
  loc_0060B76D: jge 0060B783h
  loc_0060B76F: mov edx, var_C4
  loc_0060B775: push 0000008Ch
  loc_0060B77A: push 0043B5D0h
  loc_0060B77F: push edx
  loc_0060B780: push eax
  loc_0060B781: call edi
  loc_0060B783: lea ebx, var_28
  loc_0060B786: mov eax, var_24
  loc_0060B789: push ebx
  loc_0060B78A: mov edx, 00000008h
  loc_0060B78F: sub esp, 00000010h
  loc_0060B792: mov var_70, edx
  loc_0060B795: mov ebx, esp
  loc_0060B797: mov ecx, 0043B848h ; "Section4"
  loc_0060B79C: mov var_68, ecx
  loc_0060B79F: mov esi, [eax]
  loc_0060B7A1: mov [ebx], edx
  loc_0060B7A3: mov edx, var_6C
  loc_0060B7A6: push eax
  loc_0060B7A7: mov var_AC, eax
  loc_0060B7AD: mov [ebx+00000004h], edx
  loc_0060B7B0: mov [ebx+00000008h], ecx
  loc_0060B7B3: mov ecx, var_64
  loc_0060B7B6: mov [ebx+0000000Ch], ecx
  loc_0060B7B9: call [esi+0000001Ch]
  loc_0060B7BC: test eax, eax
  loc_0060B7BE: fnclex
  loc_0060B7C0: jge 0060B7D3h
  loc_0060B7C2: mov edx, var_AC
  loc_0060B7C8: push 0000001Ch
  loc_0060B7CA: push 00439DD8h
  loc_0060B7CF: push edx
  loc_0060B7D0: push eax
  loc_0060B7D1: call edi
  loc_0060B7D3: mov eax, var_28
  loc_0060B7D6: lea edx, var_2C
  loc_0060B7D9: push edx
  loc_0060B7DA: push eax
  loc_0060B7DB: mov ecx, [eax]
  loc_0060B7DD: mov esi, eax
  loc_0060B7DF: call [ecx+0000003Ch]
  loc_0060B7E2: test eax, eax
  loc_0060B7E4: fnclex
  loc_0060B7E6: jge 0060B7F3h
  loc_0060B7E8: push 0000003Ch
  loc_0060B7EA: push 00439BF8h
  loc_0060B7EF: push esi
  loc_0060B7F0: push eax
  loc_0060B7F1: call edi
  loc_0060B7F3: lea ebx, var_30
  loc_0060B7F6: mov eax, var_2C
  loc_0060B7F9: push ebx
  loc_0060B7FA: mov edx, 00000008h
  loc_0060B7FF: sub esp, 00000010h
  loc_0060B802: mov esi, [eax]
  loc_0060B804: mov ebx, esp
  loc_0060B806: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060B80B: push eax
  loc_0060B80C: mov var_BC, eax
  loc_0060B812: mov [ebx], edx
  loc_0060B814: mov edx, var_7C
  loc_0060B817: mov [ebx+00000004h], edx
  loc_0060B81A: mov [ebx+00000008h], ecx
  loc_0060B81D: mov ecx, var_74
  loc_0060B820: mov [ebx+0000000Ch], ecx
  loc_0060B823: call [esi+0000001Ch]
  loc_0060B826: test eax, eax
  loc_0060B828: fnclex
  loc_0060B82A: jge 0060B83Dh
  loc_0060B82C: mov edx, var_BC
  loc_0060B832: push 0000001Ch
  loc_0060B834: push 0043B7D8h
  loc_0060B839: push edx
  loc_0060B83A: push eax
  loc_0060B83B: call edi
  loc_0060B83D: mov ecx, [00668108h]
  loc_0060B843: mov edx, [0066810Ch]
  loc_0060B849: sub esp, 00000010h
  loc_0060B84C: mov eax, esp
  loc_0060B84E: push 0043B85Ch ; "Caption"
  loc_0060B853: mov [eax], ecx
  loc_0060B855: mov ecx, [00668110h]
  loc_0060B85B: mov [eax+00000004h], edx
  loc_0060B85E: mov edx, [00668114h]
  loc_0060B864: mov [eax+00000008h], ecx
  loc_0060B867: mov [eax+0000000Ch], edx
  loc_0060B86A: mov eax, var_30
  loc_0060B86D: push eax
  loc_0060B86E: call [00401094h] ; __vbaLateMemSt
  loc_0060B874: lea ecx, var_30
  loc_0060B877: lea edx, var_2C
  loc_0060B87A: push ecx
  loc_0060B87B: lea eax, var_28
  loc_0060B87E: push edx
  loc_0060B87F: lea ecx, var_24
  loc_0060B882: push eax
  loc_0060B883: push ecx
  loc_0060B884: push 00000004h
  loc_0060B886: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B88C: mov eax, var_C4
  loc_0060B892: add esp, 00000014h
  loc_0060B895: mov edx, [eax]
  loc_0060B897: push 0000000Ah
  loc_0060B899: push eax
  loc_0060B89A: call [edx+00000048h]
  loc_0060B89D: test eax, eax
  loc_0060B89F: fnclex
  loc_0060B8A1: jge 0060B8B4h
  loc_0060B8A3: mov ecx, var_C4
  loc_0060B8A9: push 00000048h
  loc_0060B8AB: push 0043B5D0h
  loc_0060B8B0: push ecx
  loc_0060B8B1: push eax
  loc_0060B8B2: call edi
  loc_0060B8B4: mov eax, var_C4
  loc_0060B8BA: push 0000000Ah
  loc_0060B8BC: push eax
  loc_0060B8BD: mov edx, [eax]
  loc_0060B8BF: call [edx+00000058h]
  loc_0060B8C2: test eax, eax
  loc_0060B8C4: fnclex
  loc_0060B8C6: jge 0060B8D9h
  loc_0060B8C8: mov ecx, var_C4
  loc_0060B8CE: push 00000058h
  loc_0060B8D0: push 0043B5D0h
  loc_0060B8D5: push ecx
  loc_0060B8D6: push eax
  loc_0060B8D7: call edi
  loc_0060B8D9: sub esp, 00000010h
  loc_0060B8DC: mov eax, 00000002h
  loc_0060B8E1: mov edx, esp
  loc_0060B8E3: mov ecx, eax
  loc_0060B8E5: mov var_70, ecx
  loc_0060B8E8: mov var_68, eax
  loc_0060B8EB: mov [edx], ecx
  loc_0060B8ED: mov ecx, var_6C
  loc_0060B8F0: push 8001000Eh
  loc_0060B8F5: mov [edx+00000004h], ecx
  loc_0060B8F8: mov ecx, var_C4
  loc_0060B8FE: push ecx
  loc_0060B8FF: mov [edx+00000008h], eax
  loc_0060B902: mov eax, var_64
  loc_0060B905: mov [edx+0000000Ch], eax
  loc_0060B908: call [00401280h] ; __vbaLateIdSt
  loc_0060B90E: mov edx, var_C4
  loc_0060B914: push 00000000h
  loc_0060B916: push 80011003h
  loc_0060B91B: push edx
  loc_0060B91C: call [0040102Ch] ; __vbaLateIdCall
  loc_0060B922: add esp, 0000000Ch
  loc_0060B925: lea eax, var_C4
  loc_0060B92B: push 00000000h
  loc_0060B92D: push eax
  loc_0060B92E: call [004010C8h] ; __vbaObjSetAddref
  loc_0060B934: push 0060B99Eh
  loc_0060B939: jmp 0060B974h
  loc_0060B93B: lea ecx, var_20
  loc_0060B93E: call [0040129Ch] ; __vbaFreeStr
  loc_0060B944: lea ecx, var_30
  loc_0060B947: lea edx, var_2C
  loc_0060B94A: push ecx
  loc_0060B94B: lea eax, var_28
  loc_0060B94E: push edx
  loc_0060B94F: lea ecx, var_24
  loc_0060B952: push eax
  loc_0060B953: push ecx
  loc_0060B954: push 00000004h
  loc_0060B956: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B95C: lea edx, var_60
  loc_0060B95F: lea eax, var_50
  loc_0060B962: push edx
  loc_0060B963: lea ecx, var_40
  loc_0060B966: push eax
  loc_0060B967: push ecx
  loc_0060B968: push 00000003h
  loc_0060B96A: call [0040103Ch] ; __vbaFreeVarList
  loc_0060B970: add esp, 00000024h
  loc_0060B973: ret
  loc_0060B974: lea edx, var_C8
  loc_0060B97A: lea eax, var_C4
  loc_0060B980: push edx
  loc_0060B981: push eax
  loc_0060B982: push 00000002h
  loc_0060B984: call [0040104Ch] ; __vbaFreeObjList
  loc_0060B98A: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060B990: add esp, 0000000Ch
  loc_0060B993: lea ecx, var_18
  loc_0060B996: call __vbaFreeStr
  loc_0060B998: lea ecx, var_1C
  loc_0060B99B: call __vbaFreeStr
  loc_0060B99D: ret
  loc_0060B99E: mov ecx, var_10
  loc_0060B9A1: pop edi
  loc_0060B9A2: pop esi
  loc_0060B9A3: mov fs:[00000000h], ecx
  loc_0060B9AA: pop ebx
  loc_0060B9AB: mov esp, ebp
  loc_0060B9AD: pop ebp
  loc_0060B9AE: ret
End Sub

Private Sub Proc_34_2_60B9C0
  loc_0060B9C0: push ebp
  loc_0060B9C1: mov ebp, esp
  loc_0060B9C3: sub esp, 00000008h
  loc_0060B9C6: push 00405696h ; __vbaExceptHandler
  loc_0060B9CB: mov eax, fs:[00000000h]
  loc_0060B9D1: push eax
  loc_0060B9D2: mov fs:[00000000h], esp
  loc_0060B9D9: sub esp, 000000D0h
  loc_0060B9DF: push ebx
  loc_0060B9E0: push esi
  loc_0060B9E1: push edi
  loc_0060B9E2: mov var_8, esp
  loc_0060B9E5: mov var_4, 004044C0h
  loc_0060B9EC: xor ebx, ebx
  loc_0060B9EE: mov var_18, ebx
  loc_0060B9F1: mov var_1C, ebx
  loc_0060B9F4: mov var_20, ebx
  loc_0060B9F7: mov var_24, ebx
  loc_0060B9FA: mov var_28, ebx
  loc_0060B9FD: mov var_2C, ebx
  loc_0060BA00: mov var_30, ebx
  loc_0060BA03: mov var_40, ebx
  loc_0060BA06: mov var_50, ebx
  loc_0060BA09: mov var_60, ebx
  loc_0060BA0C: mov var_70, ebx
  loc_0060BA0F: mov var_80, ebx
  loc_0060BA12: mov var_90, ebx
  loc_0060BA18: mov var_A4, ebx
  loc_0060BA1E: mov var_C4, ebx
  loc_0060BA24: mov var_C8, ebx
  loc_0060BA2A: call 0055B320h
  loc_0060BA2F: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0060BA35: mov edx, 0042DFECh
  loc_0060BA3A: lea ecx, var_18
  loc_0060BA3D: call __vbaStrCopy
  loc_0060BA3F: mov edx, 0043B9C0h ; " SELECT * FROM pelanggan ORDER BY kode_pelanggan"
  loc_0060BA44: lea ecx, var_18
  loc_0060BA47: call __vbaStrCopy
  loc_0060BA49: push 0042DF30h
  loc_0060BA4E: call [00401168h] ; __vbaNew
  loc_0060BA54: push eax
  loc_0060BA55: push 00668090h
  loc_0060BA5A: call [004010B8h] ; __vbaObjSet
  loc_0060BA60: cmp [0066803Ch], ebx
  loc_0060BA66: jnz 0060BA78h
  loc_0060BA68: push 0066803Ch
  loc_0060BA6D: push 0042DEFCh
  loc_0060BA72: call [004011E8h] ; __vbaNew2
  loc_0060BA78: mov esi, [0066803Ch]
  loc_0060BA7E: lea ecx, var_40
  loc_0060BA81: mov var_38, 80020004h
  loc_0060BA88: mov var_40, 0000000Ah
  loc_0060BA8F: call [0040123Ch] ; __vbaFreeVarg
  loc_0060BA95: mov eax, [esi]
  loc_0060BA97: lea ecx, var_24
  loc_0060BA9A: push ecx
  loc_0060BA9B: mov ecx, var_18
  loc_0060BA9E: lea edx, var_40
  loc_0060BAA1: push 00000001h
  loc_0060BAA3: push edx
  loc_0060BAA4: push ecx
  loc_0060BAA5: push esi
  loc_0060BAA6: call [eax+00000040h]
  loc_0060BAA9: cmp eax, ebx
  loc_0060BAAB: fnclex
  loc_0060BAAD: jge 0060BAC2h
  loc_0060BAAF: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060BAB5: push 00000040h
  loc_0060BAB7: push 0042E1B0h
  loc_0060BABC: push esi
  loc_0060BABD: push eax
  loc_0060BABE: call edi
  loc_0060BAC0: jmp 0060BAC8h
  loc_0060BAC2: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060BAC8: mov edx, var_24
  loc_0060BACB: mov esi, [00401268h] ; __vbaCastObj
  loc_0060BAD1: push 0042DF20h
  loc_0060BAD6: push edx
  loc_0060BAD7: call __vbaCastObj
  loc_0060BAD9: push eax
  loc_0060BADA: push 00668090h
  loc_0060BADF: call [004010B8h] ; __vbaObjSet
  loc_0060BAE5: lea ecx, var_24
  loc_0060BAE8: call [004012A0h] ; __vbaFreeObj
  loc_0060BAEE: lea ecx, var_40
  loc_0060BAF1: call [00401020h] ; __vbaFreeVar
  loc_0060BAF7: cmp [00668530h], ebx
  loc_0060BAFD: jnz 0060BB0Fh
  loc_0060BAFF: push 00668530h
  loc_0060BB04: push 00407610h
  loc_0060BB09: call [004011E8h] ; __vbaNew2
  loc_0060BB0F: mov eax, [00668530h]
  loc_0060BB14: lea ecx, var_C4
  loc_0060BB1A: push eax
  loc_0060BB1B: push ecx
  loc_0060BB1C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060BB22: mov edx, var_C4
  loc_0060BB28: push 0043B7B8h
  loc_0060BB2D: push ebx
  loc_0060BB2E: mov edx, [edx]
  loc_0060BB30: mov var_E0, edx
  loc_0060BB36: call __vbaCastObj
  loc_0060BB38: push eax
  loc_0060BB39: lea eax, var_24
  loc_0060BB3C: push eax
  loc_0060BB3D: call [004010B8h] ; __vbaObjSet
  loc_0060BB43: mov ecx, var_C4
  loc_0060BB49: mov edx, var_E0
  loc_0060BB4F: push eax
  loc_0060BB50: push ecx
  loc_0060BB51: call [edx+00000028h]
  loc_0060BB54: cmp eax, ebx
  loc_0060BB56: fnclex
  loc_0060BB58: jge 0060BB6Bh
  loc_0060BB5A: mov ecx, var_C4
  loc_0060BB60: push 00000028h
  loc_0060BB62: push 0043B5D0h
  loc_0060BB67: push ecx
  loc_0060BB68: push eax
  loc_0060BB69: call edi
  loc_0060BB6B: lea ecx, var_24
  loc_0060BB6E: call [004012A0h] ; __vbaFreeObj
  loc_0060BB74: mov eax, var_C4
  loc_0060BB7A: push 0042DFECh
  loc_0060BB7F: push eax
  loc_0060BB80: mov edx, [eax]
  loc_0060BB82: call [edx+00000030h]
  loc_0060BB85: cmp eax, ebx
  loc_0060BB87: fnclex
  loc_0060BB89: jge 0060BB9Ch
  loc_0060BB8B: mov ecx, var_C4
  loc_0060BB91: push 00000030h
  loc_0060BB93: push 0043B5D0h
  loc_0060BB98: push ecx
  loc_0060BB99: push eax
  loc_0060BB9A: call edi
  loc_0060BB9C: mov eax, [00668090h]
  loc_0060BBA1: lea ecx, var_24
  loc_0060BBA4: push ecx
  loc_0060BBA5: push eax
  loc_0060BBA6: mov edx, [eax]
  loc_0060BBA8: call [edx+00000114h]
  loc_0060BBAE: cmp eax, ebx
  loc_0060BBB0: fnclex
  loc_0060BBB2: jge 0060BBC8h
  loc_0060BBB4: mov edx, [00668090h]
  loc_0060BBBA: push 00000114h
  loc_0060BBBF: push 0043B7C8h
  loc_0060BBC4: push edx
  loc_0060BBC5: push eax
  loc_0060BBC6: call edi
  loc_0060BBC8: mov eax, var_C4
  loc_0060BBCE: mov ecx, var_24
  loc_0060BBD1: push 0043B7B8h
  loc_0060BBD6: push ecx
  loc_0060BBD7: mov esi, [eax]
  loc_0060BBD9: call [00401268h] ; __vbaCastObj
  loc_0060BBDF: lea edx, var_28
  loc_0060BBE2: push eax
  loc_0060BBE3: push edx
  loc_0060BBE4: call [004010B8h] ; __vbaObjSet
  loc_0060BBEA: push eax
  loc_0060BBEB: mov eax, var_C4
  loc_0060BBF1: push eax
  loc_0060BBF2: call [esi+00000028h]
  loc_0060BBF5: cmp eax, ebx
  loc_0060BBF7: fnclex
  loc_0060BBF9: jge 0060BC0Ch
  loc_0060BBFB: mov ecx, var_C4
  loc_0060BC01: push 00000028h
  loc_0060BC03: push 0043B5D0h
  loc_0060BC08: push ecx
  loc_0060BC09: push eax
  loc_0060BC0A: call edi
  loc_0060BC0C: lea edx, var_28
  loc_0060BC0F: lea eax, var_24
  loc_0060BC12: push edx
  loc_0060BC13: push eax
  loc_0060BC14: push 00000002h
  loc_0060BC16: call [0040104Ch] ; __vbaFreeObjList
  loc_0060BC1C: mov eax, var_C4
  loc_0060BC22: add esp, 0000000Ch
  loc_0060BC25: lea edx, var_24
  loc_0060BC28: mov ecx, [eax]
  loc_0060BC2A: push edx
  loc_0060BC2B: push eax
  loc_0060BC2C: call [ecx+0000008Ch]
  loc_0060BC32: cmp eax, ebx
  loc_0060BC34: fnclex
  loc_0060BC36: jge 0060BC4Ch
  loc_0060BC38: mov ecx, var_C4
  loc_0060BC3E: push 0000008Ch
  loc_0060BC43: push 0043B5D0h
  loc_0060BC48: push ecx
  loc_0060BC49: push eax
  loc_0060BC4A: call edi
  loc_0060BC4C: lea esi, var_28
  loc_0060BC4F: mov eax, var_24
  loc_0060BC52: push esi
  loc_0060BC53: mov ecx, 00000008h
  loc_0060BC58: sub esp, 00000010h
  loc_0060BC5B: mov var_70, ecx
  loc_0060BC5E: mov esi, esp
  loc_0060BC60: mov var_68, 00439FB0h ; "Section1"
  loc_0060BC67: mov edx, [eax]
  loc_0060BC69: push eax
  loc_0060BC6A: mov [esi], ecx
  loc_0060BC6C: mov ecx, var_6C
  loc_0060BC6F: mov var_AC, eax
  loc_0060BC75: mov [esi+00000004h], ecx
  loc_0060BC78: mov ecx, var_68
  loc_0060BC7B: mov [esi+00000008h], ecx
  loc_0060BC7E: mov ecx, var_64
  loc_0060BC81: mov [esi+0000000Ch], ecx
  loc_0060BC84: call [edx+0000001Ch]
  loc_0060BC87: cmp eax, ebx
  loc_0060BC89: fnclex
  loc_0060BC8B: jge 0060BC9Eh
  loc_0060BC8D: mov edx, var_AC
  loc_0060BC93: push 0000001Ch
  loc_0060BC95: push 00439DD8h
  loc_0060BC9A: push edx
  loc_0060BC9B: push eax
  loc_0060BC9C: call edi
  loc_0060BC9E: mov eax, var_28
  loc_0060BCA1: lea edx, var_2C
  loc_0060BCA4: push edx
  loc_0060BCA5: push eax
  loc_0060BCA6: mov ecx, [eax]
  loc_0060BCA8: mov esi, eax
  loc_0060BCAA: call [ecx+0000003Ch]
  loc_0060BCAD: cmp eax, ebx
  loc_0060BCAF: fnclex
  loc_0060BCB1: jge 0060BCBEh
  loc_0060BCB3: push 0000003Ch
  loc_0060BCB5: push 00439BF8h
  loc_0060BCBA: push esi
  loc_0060BCBB: push eax
  loc_0060BCBC: call edi
  loc_0060BCBE: mov eax, var_2C
  loc_0060BCC1: mov var_2C, ebx
  loc_0060BCC4: push eax
  loc_0060BCC5: lea eax, var_C8
  loc_0060BCCB: push eax
  loc_0060BCCC: call [004010B8h] ; __vbaObjSet
  loc_0060BCD2: lea ecx, var_28
  loc_0060BCD5: lea edx, var_24
  loc_0060BCD8: push ecx
  loc_0060BCD9: push edx
  loc_0060BCDA: push 00000002h
  loc_0060BCDC: call [0040104Ch] ; __vbaFreeObjList
  loc_0060BCE2: mov eax, var_C8
  loc_0060BCE8: add esp, 0000000Ch
  loc_0060BCEB: lea edx, var_A4
  loc_0060BCF1: mov ecx, [eax]
  loc_0060BCF3: push edx
  loc_0060BCF4: push eax
  loc_0060BCF5: call [ecx+00000020h]
  loc_0060BCF8: cmp eax, ebx
  loc_0060BCFA: fnclex
  loc_0060BCFC: jge 0060BD0Fh
  loc_0060BCFE: mov ecx, var_C8
  loc_0060BD04: push 00000020h
  loc_0060BD06: push 0043B7D8h
  loc_0060BD0B: push ecx
  loc_0060BD0C: push eax
  loc_0060BD0D: call edi
  loc_0060BD0F: mov ecx, var_A4
  loc_0060BD15: call [00401138h] ; __vbaI2I4
  loc_0060BD1B: mov var_D0, eax
  loc_0060BD21: mov ebx, 00000001h
  loc_0060BD26: cmp bx, var_D0
  loc_0060BD2D: mov var_14, ebx
  loc_0060BD30: jg 0060BFA1h
  loc_0060BD36: lea esi, var_24
  loc_0060BD39: mov ecx, var_C8
  loc_0060BD3F: push esi
  loc_0060BD40: mov eax, 00000002h
  loc_0060BD45: sub esp, 00000010h
  loc_0060BD48: mov var_70, eax
  loc_0060BD4B: mov esi, esp
  loc_0060BD4D: mov var_68, bx
  loc_0060BD51: mov edx, [ecx]
  loc_0060BD53: push ecx
  loc_0060BD54: mov [esi], eax
  loc_0060BD56: mov eax, var_6C
  loc_0060BD59: mov [esi+00000004h], eax
  loc_0060BD5C: mov eax, var_68
  loc_0060BD5F: mov [esi+00000008h], eax
  loc_0060BD62: mov eax, var_64
  loc_0060BD65: mov [esi+0000000Ch], eax
  loc_0060BD68: call [edx+0000001Ch]
  loc_0060BD6B: test eax, eax
  loc_0060BD6D: fnclex
  loc_0060BD6F: jge 0060BD82h
  loc_0060BD71: mov ecx, var_C8
  loc_0060BD77: push 0000001Ch
  loc_0060BD79: push 0043B7D8h
  loc_0060BD7E: push ecx
  loc_0060BD7F: push eax
  loc_0060BD80: call edi
  loc_0060BD82: mov edx, var_24
  loc_0060BD85: push 0043B7E8h
  loc_0060BD8A: push edx
  loc_0060BD8B: call [004011CCh] ; __vbaCheckType
  loc_0060BD91: lea ecx, var_24
  loc_0060BD94: mov si, ax
  loc_0060BD97: call [004012A0h] ; __vbaFreeObj
  loc_0060BD9D: test si, si
  loc_0060BDA0: Unknown_795()
  loc_0060BDA6: mov var_68, bx
  loc_0060BDAA: lea ebx, var_24
  loc_0060BDAD: push ebx
  loc_0060BDAE: mov ecx, var_C8
  loc_0060BDB4: sub esp, 00000010h
  loc_0060BDB7: mov eax, 00000002h
  loc_0060BDBC: mov ebx, esp
  loc_0060BDBE: mov var_70, eax
  loc_0060BDC1: mov edx, [ecx]
  loc_0060BDC3: mov esi, 0042DFECh
  loc_0060BDC8: mov [ebx], eax
  loc_0060BDCA: mov eax, var_6C
  loc_0060BDCD: push ecx
  loc_0060BDCE: mov var_78, esi
  loc_0060BDD1: mov [ebx+00000004h], eax
  loc_0060BDD4: mov eax, var_68
  loc_0060BDD7: mov edi, 00000008h
  loc_0060BDDC: mov [ebx+00000008h], eax
  loc_0060BDDF: mov eax, var_64
  loc_0060BDE2: mov [ebx+0000000Ch], eax
  loc_0060BDE5: call [edx+0000001Ch]
  loc_0060BDE8: test eax, eax
  loc_0060BDEA: fnclex
  loc_0060BDEC: jge 0060BE03h
  loc_0060BDEE: mov ecx, var_C8
  loc_0060BDF4: push 0000001Ch
  loc_0060BDF6: push 0043B7D8h
  loc_0060BDFB: push ecx
  loc_0060BDFC: push eax
  loc_0060BDFD: call [00401080h] ; __vbaHresultCheckObj
  loc_0060BE03: mov eax, var_7C
  loc_0060BE06: sub esp, 00000010h
  loc_0060BE09: mov edx, esp
  loc_0060BE0B: mov ecx, var_74
  loc_0060BE0E: push 0043B7F8h ; "DataMember"
  loc_0060BE13: mov [edx], edi
  loc_0060BE15: mov [edx+00000004h], eax
  loc_0060BE18: mov [edx+00000008h], esi
  loc_0060BE1B: mov [edx+0000000Ch], ecx
  loc_0060BE1E: mov edx, var_24
  loc_0060BE21: push edx
  loc_0060BE22: call [00401094h] ; __vbaLateMemSt
  loc_0060BE28: lea ecx, var_24
  loc_0060BE2B: call [004012A0h] ; __vbaFreeObj
  loc_0060BE31: mov eax, [00668090h]
  loc_0060BE36: lea edx, var_24
  loc_0060BE39: push edx
  loc_0060BE3A: push eax
  loc_0060BE3B: mov ecx, [eax]
  loc_0060BE3D: call [ecx+00000054h]
  loc_0060BE40: test eax, eax
  loc_0060BE42: fnclex
  loc_0060BE44: jge 0060BE5Bh
  loc_0060BE46: mov ecx, [00668090h]
  loc_0060BE4C: push 00000054h
  loc_0060BE4E: push 0042DF44h
  loc_0060BE53: push ecx
  loc_0060BE54: push eax
  loc_0060BE55: call [00401080h] ; __vbaHresultCheckObj
  loc_0060BE5B: mov ebx, var_14
  loc_0060BE5E: lea edi, var_28
  loc_0060BE61: mov dx, bx
  loc_0060BE64: push edi
  loc_0060BE65: sub dx, 0001h
  loc_0060BE69: mov eax, var_24
  loc_0060BE6C: jo 0060C72Fh
  loc_0060BE72: sub esp, 00000010h
  loc_0060BE75: mov ecx, 00000002h
  loc_0060BE7A: mov edi, esp
  loc_0060BE7C: mov var_70, ecx
  loc_0060BE7F: mov var_68, dx
  loc_0060BE83: mov edx, [eax]
  loc_0060BE85: mov [edi], ecx
  loc_0060BE87: mov ecx, var_6C
  loc_0060BE8A: push eax
  loc_0060BE8B: mov esi, eax
  loc_0060BE8D: mov [edi+00000004h], ecx
  loc_0060BE90: mov ecx, var_68
  loc_0060BE93: mov [edi+00000008h], ecx
  loc_0060BE96: mov ecx, var_64
  loc_0060BE99: mov [edi+0000000Ch], ecx
  loc_0060BE9C: call [edx+00000028h]
  loc_0060BE9F: test eax, eax
  loc_0060BEA1: fnclex
  loc_0060BEA3: jge 0060BEB8h
  loc_0060BEA5: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060BEAB: push 00000028h
  loc_0060BEAD: push 0042DFACh
  loc_0060BEB2: push esi
  loc_0060BEB3: push eax
  loc_0060BEB4: call edi
  loc_0060BEB6: jmp 0060BEBEh
  loc_0060BEB8: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060BEBE: mov eax, var_28
  loc_0060BEC1: lea ecx, var_20
  loc_0060BEC4: push ecx
  loc_0060BEC5: push eax
  loc_0060BEC6: mov edx, [eax]
  loc_0060BEC8: mov esi, eax
  loc_0060BECA: call [edx+0000002Ch]
  loc_0060BECD: test eax, eax
  loc_0060BECF: fnclex
  loc_0060BED1: jge 0060BEDEh
  loc_0060BED3: push 0000002Ch
  loc_0060BED5: push 0042DFBCh
  loc_0060BEDA: push esi
  loc_0060BEDB: push eax
  loc_0060BEDC: call edi
  loc_0060BEDE: mov eax, var_20
  loc_0060BEE1: lea esi, var_2C
  loc_0060BEE4: push esi
  loc_0060BEE5: mov ecx, var_C8
  loc_0060BEEB: sub esp, 00000010h
  loc_0060BEEE: mov var_38, eax
  loc_0060BEF1: mov esi, esp
  loc_0060BEF3: mov eax, 00000002h
  loc_0060BEF8: mov var_78, bx
  loc_0060BEFC: mov var_20, 00000000h
  loc_0060BF03: mov [esi], eax
  loc_0060BF05: mov eax, var_7C
  loc_0060BF08: mov var_40, 00000008h
  loc_0060BF0F: mov edx, [ecx]
  loc_0060BF11: mov [esi+00000004h], eax
  loc_0060BF14: mov eax, var_78
  loc_0060BF17: push ecx
  loc_0060BF18: mov [esi+00000008h], eax
  loc_0060BF1B: mov eax, var_74
  loc_0060BF1E: mov [esi+0000000Ch], eax
  loc_0060BF21: call [edx+0000001Ch]
  loc_0060BF24: test eax, eax
  loc_0060BF26: fnclex
  loc_0060BF28: jge 0060BF3Bh
  loc_0060BF2A: mov ecx, var_C8
  loc_0060BF30: push 0000001Ch
  loc_0060BF32: push 0043B7D8h
  loc_0060BF37: push ecx
  loc_0060BF38: push eax
  loc_0060BF39: call edi
  loc_0060BF3B: mov eax, var_40
  loc_0060BF3E: mov ecx, var_3C
  loc_0060BF41: sub esp, 00000010h
  loc_0060BF44: mov edx, esp
  loc_0060BF46: push 0043B810h ; "DataField"
  loc_0060BF4B: mov [edx], eax
  loc_0060BF4D: mov eax, var_38
  loc_0060BF50: mov [edx+00000004h], ecx
  loc_0060BF53: mov ecx, var_34
  loc_0060BF56: mov [edx+00000008h], eax
  loc_0060BF59: mov [edx+0000000Ch], ecx
  loc_0060BF5C: mov edx, var_2C
  loc_0060BF5F: push edx
  loc_0060BF60: call [00401094h] ; __vbaLateMemSt
  loc_0060BF66: lea eax, var_2C
  loc_0060BF69: lea ecx, var_28
  loc_0060BF6C: push eax
  loc_0060BF6D: lea edx, var_24
  loc_0060BF70: push ecx
  loc_0060BF71: push edx
  loc_0060BF72: push 00000003h
  loc_0060BF74: call [0040104Ch] ; __vbaFreeObjList
  loc_0060BF7A: add esp, 00000010h
  loc_0060BF7D: lea ecx, var_40
  loc_0060BF80: call [00401020h] ; __vbaFreeVar
  loc_0060BF86: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060BF8C: mov eax, 00000001h
  loc_0060BF91: add ax, bx
  loc_0060BF94: jo 0060C72Fh
  loc_0060BF9A: mov ebx, eax
  loc_0060BF9C: jmp 0060BD26h
  loc_0060BFA1: lea eax, var_C8
  loc_0060BFA7: push 00000000h
  loc_0060BFA9: push eax
  loc_0060BFAA: call [004010C8h] ; __vbaObjSetAddref
  loc_0060BFB0: lea ecx, var_40
  loc_0060BFB3: push ecx
  loc_0060BFB4: call [00401220h] ; rtcGetDateVar
  loc_0060BFBA: lea edx, var_70
  loc_0060BFBD: lea ecx, var_50
  loc_0060BFC0: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060BFC7: mov var_70, 00000008h
  loc_0060BFCE: call [00401238h] ; __vbaVarDup
  loc_0060BFD4: push 00000001h
  loc_0060BFD6: lea edx, var_50
  loc_0060BFD9: push 00000001h
  loc_0060BFDB: lea eax, var_40
  loc_0060BFDE: push edx
  loc_0060BFDF: lea ecx, var_60
  loc_0060BFE2: push eax
  loc_0060BFE3: push ecx
  loc_0060BFE4: call [00401078h] ; rtcVarFromFormatVar
  loc_0060BFEA: mov eax, var_C4
  loc_0060BFF0: lea ecx, var_24
  loc_0060BFF3: push ecx
  loc_0060BFF4: push eax
  loc_0060BFF5: mov edx, [eax]
  loc_0060BFF7: call [edx+0000008Ch]
  loc_0060BFFD: test eax, eax
  loc_0060BFFF: fnclex
  loc_0060C001: jge 0060C017h
  loc_0060C003: mov edx, var_C4
  loc_0060C009: push 0000008Ch
  loc_0060C00E: push 0043B5D0h
  loc_0060C013: push edx
  loc_0060C014: push eax
  loc_0060C015: call edi
  loc_0060C017: lea ebx, var_28
  loc_0060C01A: mov eax, var_24
  loc_0060C01D: push ebx
  loc_0060C01E: mov edx, 00000008h
  loc_0060C023: sub esp, 00000010h
  loc_0060C026: mov esi, [eax]
  loc_0060C028: mov ebx, esp
  loc_0060C02A: mov ecx, 0043B848h ; "Section4"
  loc_0060C02F: push eax
  loc_0060C030: mov var_AC, eax
  loc_0060C036: mov [ebx], edx
  loc_0060C038: mov edx, var_7C
  loc_0060C03B: mov [ebx+00000004h], edx
  loc_0060C03E: mov [ebx+00000008h], ecx
  loc_0060C041: mov ecx, var_74
  loc_0060C044: mov [ebx+0000000Ch], ecx
  loc_0060C047: call [esi+0000001Ch]
  loc_0060C04A: test eax, eax
  loc_0060C04C: fnclex
  loc_0060C04E: jge 0060C061h
  loc_0060C050: mov edx, var_AC
  loc_0060C056: push 0000001Ch
  loc_0060C058: push 00439DD8h
  loc_0060C05D: push edx
  loc_0060C05E: push eax
  loc_0060C05F: call edi
  loc_0060C061: mov eax, var_28
  loc_0060C064: lea edx, var_2C
  loc_0060C067: push edx
  loc_0060C068: push eax
  loc_0060C069: mov ecx, [eax]
  loc_0060C06B: mov esi, eax
  loc_0060C06D: call [ecx+0000003Ch]
  loc_0060C070: test eax, eax
  loc_0060C072: fnclex
  loc_0060C074: jge 0060C081h
  loc_0060C076: push 0000003Ch
  loc_0060C078: push 00439BF8h
  loc_0060C07D: push esi
  loc_0060C07E: push eax
  loc_0060C07F: call edi
  loc_0060C081: lea ebx, var_30
  loc_0060C084: mov eax, var_2C
  loc_0060C087: push ebx
  loc_0060C088: mov edx, 00000008h
  loc_0060C08D: sub esp, 00000010h
  loc_0060C090: mov esi, [eax]
  loc_0060C092: mov ebx, esp
  loc_0060C094: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060C099: push eax
  loc_0060C09A: mov var_BC, eax
  loc_0060C0A0: mov [ebx], edx
  loc_0060C0A2: mov edx, var_8C
  loc_0060C0A8: mov [ebx+00000004h], edx
  loc_0060C0AB: mov [ebx+00000008h], ecx
  loc_0060C0AE: mov ecx, var_84
  loc_0060C0B4: mov [ebx+0000000Ch], ecx
  loc_0060C0B7: call [esi+0000001Ch]
  loc_0060C0BA: test eax, eax
  loc_0060C0BC: fnclex
  loc_0060C0BE: jge 0060C0D1h
  loc_0060C0C0: mov edx, var_BC
  loc_0060C0C6: push 0000001Ch
  loc_0060C0C8: push 0043B7D8h
  loc_0060C0CD: push edx
  loc_0060C0CE: push eax
  loc_0060C0CF: call edi
  loc_0060C0D1: mov ecx, var_60
  loc_0060C0D4: mov edx, var_5C
  loc_0060C0D7: sub esp, 00000010h
  loc_0060C0DA: mov eax, esp
  loc_0060C0DC: push 0043B85Ch ; "Caption"
  loc_0060C0E1: mov [eax], ecx
  loc_0060C0E3: mov ecx, var_58
  loc_0060C0E6: mov [eax+00000004h], edx
  loc_0060C0E9: mov edx, var_54
  loc_0060C0EC: mov [eax+00000008h], ecx
  loc_0060C0EF: mov [eax+0000000Ch], edx
  loc_0060C0F2: mov eax, var_30
  loc_0060C0F5: push eax
  loc_0060C0F6: call [00401094h] ; __vbaLateMemSt
  loc_0060C0FC: lea ecx, var_30
  loc_0060C0FF: lea edx, var_2C
  loc_0060C102: push ecx
  loc_0060C103: lea eax, var_28
  loc_0060C106: push edx
  loc_0060C107: lea ecx, var_24
  loc_0060C10A: push eax
  loc_0060C10B: push ecx
  loc_0060C10C: push 00000004h
  loc_0060C10E: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C114: lea edx, var_60
  loc_0060C117: lea eax, var_50
  loc_0060C11A: push edx
  loc_0060C11B: lea ecx, var_40
  loc_0060C11E: push eax
  loc_0060C11F: push ecx
  loc_0060C120: push 00000003h
  loc_0060C122: call [0040103Ch] ; __vbaFreeVarList
  loc_0060C128: mov eax, var_C4
  loc_0060C12E: add esp, 00000024h
  loc_0060C131: lea ecx, var_24
  loc_0060C134: mov edx, [eax]
  loc_0060C136: push ecx
  loc_0060C137: push eax
  loc_0060C138: call [edx+0000008Ch]
  loc_0060C13E: test eax, eax
  loc_0060C140: fnclex
  loc_0060C142: jge 0060C158h
  loc_0060C144: mov edx, var_C4
  loc_0060C14A: push 0000008Ch
  loc_0060C14F: push 0043B5D0h
  loc_0060C154: push edx
  loc_0060C155: push eax
  loc_0060C156: call edi
  loc_0060C158: lea ebx, var_28
  loc_0060C15B: mov eax, var_24
  loc_0060C15E: push ebx
  loc_0060C15F: mov edx, 00000008h
  loc_0060C164: sub esp, 00000010h
  loc_0060C167: mov var_70, edx
  loc_0060C16A: mov ebx, esp
  loc_0060C16C: mov ecx, 0043B848h ; "Section4"
  loc_0060C171: mov var_68, ecx
  loc_0060C174: mov esi, [eax]
  loc_0060C176: mov [ebx], edx
  loc_0060C178: mov edx, var_6C
  loc_0060C17B: push eax
  loc_0060C17C: mov var_AC, eax
  loc_0060C182: mov [ebx+00000004h], edx
  loc_0060C185: mov [ebx+00000008h], ecx
  loc_0060C188: mov ecx, var_64
  loc_0060C18B: mov [ebx+0000000Ch], ecx
  loc_0060C18E: call [esi+0000001Ch]
  loc_0060C191: test eax, eax
  loc_0060C193: fnclex
  loc_0060C195: jge 0060C1A8h
  loc_0060C197: mov edx, var_AC
  loc_0060C19D: push 0000001Ch
  loc_0060C19F: push 00439DD8h
  loc_0060C1A4: push edx
  loc_0060C1A5: push eax
  loc_0060C1A6: call edi
  loc_0060C1A8: mov eax, var_28
  loc_0060C1AB: lea edx, var_2C
  loc_0060C1AE: push edx
  loc_0060C1AF: push eax
  loc_0060C1B0: mov ecx, [eax]
  loc_0060C1B2: mov esi, eax
  loc_0060C1B4: call [ecx+0000003Ch]
  loc_0060C1B7: test eax, eax
  loc_0060C1B9: fnclex
  loc_0060C1BB: jge 0060C1C8h
  loc_0060C1BD: push 0000003Ch
  loc_0060C1BF: push 00439BF8h
  loc_0060C1C4: push esi
  loc_0060C1C5: push eax
  loc_0060C1C6: call edi
  loc_0060C1C8: lea ebx, var_30
  loc_0060C1CB: mov eax, var_2C
  loc_0060C1CE: push ebx
  loc_0060C1CF: mov edx, 00000008h
  loc_0060C1D4: sub esp, 00000010h
  loc_0060C1D7: mov esi, [eax]
  loc_0060C1D9: mov ebx, esp
  loc_0060C1DB: mov ecx, 0043B870h ; "LabelNama"
  loc_0060C1E0: push eax
  loc_0060C1E1: mov var_BC, eax
  loc_0060C1E7: mov [ebx], edx
  loc_0060C1E9: mov edx, var_7C
  loc_0060C1EC: mov [ebx+00000004h], edx
  loc_0060C1EF: mov [ebx+00000008h], ecx
  loc_0060C1F2: mov ecx, var_74
  loc_0060C1F5: mov [ebx+0000000Ch], ecx
  loc_0060C1F8: call [esi+0000001Ch]
  loc_0060C1FB: test eax, eax
  loc_0060C1FD: fnclex
  loc_0060C1FF: jge 0060C212h
  loc_0060C201: mov edx, var_BC
  loc_0060C207: push 0000001Ch
  loc_0060C209: push 0043B7D8h
  loc_0060C20E: push edx
  loc_0060C20F: push eax
  loc_0060C210: call edi
  loc_0060C212: mov ecx, [006680D8h]
  loc_0060C218: mov edx, [006680DCh]
  loc_0060C21E: sub esp, 00000010h
  loc_0060C221: mov eax, esp
  loc_0060C223: push 0043B85Ch ; "Caption"
  loc_0060C228: mov [eax], ecx
  loc_0060C22A: mov ecx, [006680E0h]
  loc_0060C230: mov [eax+00000004h], edx
  loc_0060C233: mov edx, [006680E4h]
  loc_0060C239: mov [eax+00000008h], ecx
  loc_0060C23C: mov [eax+0000000Ch], edx
  loc_0060C23F: mov eax, var_30
  loc_0060C242: push eax
  loc_0060C243: call [00401094h] ; __vbaLateMemSt
  loc_0060C249: lea ecx, var_30
  loc_0060C24C: lea edx, var_2C
  loc_0060C24F: push ecx
  loc_0060C250: lea eax, var_28
  loc_0060C253: push edx
  loc_0060C254: lea ecx, var_24
  loc_0060C257: push eax
  loc_0060C258: push ecx
  loc_0060C259: push 00000004h
  loc_0060C25B: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C261: mov eax, var_C4
  loc_0060C267: add esp, 00000014h
  loc_0060C26A: lea ecx, var_24
  loc_0060C26D: mov edx, [eax]
  loc_0060C26F: push ecx
  loc_0060C270: push eax
  loc_0060C271: call [edx+0000008Ch]
  loc_0060C277: test eax, eax
  loc_0060C279: fnclex
  loc_0060C27B: jge 0060C291h
  loc_0060C27D: mov edx, var_C4
  loc_0060C283: push 0000008Ch
  loc_0060C288: push 0043B5D0h
  loc_0060C28D: push edx
  loc_0060C28E: push eax
  loc_0060C28F: call edi
  loc_0060C291: lea ebx, var_28
  loc_0060C294: mov eax, var_24
  loc_0060C297: push ebx
  loc_0060C298: mov edx, 00000008h
  loc_0060C29D: sub esp, 00000010h
  loc_0060C2A0: mov var_70, edx
  loc_0060C2A3: mov ebx, esp
  loc_0060C2A5: mov ecx, 0043B848h ; "Section4"
  loc_0060C2AA: mov var_68, ecx
  loc_0060C2AD: mov esi, [eax]
  loc_0060C2AF: mov [ebx], edx
  loc_0060C2B1: mov edx, var_6C
  loc_0060C2B4: push eax
  loc_0060C2B5: mov var_AC, eax
  loc_0060C2BB: mov [ebx+00000004h], edx
  loc_0060C2BE: mov [ebx+00000008h], ecx
  loc_0060C2C1: mov ecx, var_64
  loc_0060C2C4: mov [ebx+0000000Ch], ecx
  loc_0060C2C7: call [esi+0000001Ch]
  loc_0060C2CA: test eax, eax
  loc_0060C2CC: fnclex
  loc_0060C2CE: jge 0060C2E1h
  loc_0060C2D0: mov edx, var_AC
  loc_0060C2D6: push 0000001Ch
  loc_0060C2D8: push 00439DD8h
  loc_0060C2DD: push edx
  loc_0060C2DE: push eax
  loc_0060C2DF: call edi
  loc_0060C2E1: mov eax, var_28
  loc_0060C2E4: lea edx, var_2C
  loc_0060C2E7: push edx
  loc_0060C2E8: push eax
  loc_0060C2E9: mov ecx, [eax]
  loc_0060C2EB: mov esi, eax
  loc_0060C2ED: call [ecx+0000003Ch]
  loc_0060C2F0: test eax, eax
  loc_0060C2F2: fnclex
  loc_0060C2F4: jge 0060C301h
  loc_0060C2F6: push 0000003Ch
  loc_0060C2F8: push 00439BF8h
  loc_0060C2FD: push esi
  loc_0060C2FE: push eax
  loc_0060C2FF: call edi
  loc_0060C301: lea ebx, var_30
  loc_0060C304: mov eax, var_2C
  loc_0060C307: push ebx
  loc_0060C308: mov edx, 00000008h
  loc_0060C30D: sub esp, 00000010h
  loc_0060C310: mov esi, [eax]
  loc_0060C312: mov ebx, esp
  loc_0060C314: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060C319: push eax
  loc_0060C31A: mov var_BC, eax
  loc_0060C320: mov [ebx], edx
  loc_0060C322: mov edx, var_7C
  loc_0060C325: mov [ebx+00000004h], edx
  loc_0060C328: mov [ebx+00000008h], ecx
  loc_0060C32B: mov ecx, var_74
  loc_0060C32E: mov [ebx+0000000Ch], ecx
  loc_0060C331: call [esi+0000001Ch]
  loc_0060C334: test eax, eax
  loc_0060C336: fnclex
  loc_0060C338: jge 0060C34Bh
  loc_0060C33A: mov edx, var_BC
  loc_0060C340: push 0000001Ch
  loc_0060C342: push 0043B7D8h
  loc_0060C347: push edx
  loc_0060C348: push eax
  loc_0060C349: call edi
  loc_0060C34B: mov ecx, [006680E8h]
  loc_0060C351: mov edx, [006680ECh]
  loc_0060C357: sub esp, 00000010h
  loc_0060C35A: mov eax, esp
  loc_0060C35C: push 0043B85Ch ; "Caption"
  loc_0060C361: mov [eax], ecx
  loc_0060C363: mov ecx, [006680F0h]
  loc_0060C369: mov [eax+00000004h], edx
  loc_0060C36C: mov edx, [006680F4h]
  loc_0060C372: mov [eax+00000008h], ecx
  loc_0060C375: mov [eax+0000000Ch], edx
  loc_0060C378: mov eax, var_30
  loc_0060C37B: push eax
  loc_0060C37C: call [00401094h] ; __vbaLateMemSt
  loc_0060C382: lea ecx, var_30
  loc_0060C385: lea edx, var_2C
  loc_0060C388: push ecx
  loc_0060C389: lea eax, var_28
  loc_0060C38C: push edx
  loc_0060C38D: lea ecx, var_24
  loc_0060C390: push eax
  loc_0060C391: push ecx
  loc_0060C392: push 00000004h
  loc_0060C394: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C39A: mov eax, var_C4
  loc_0060C3A0: add esp, 00000014h
  loc_0060C3A3: lea ecx, var_24
  loc_0060C3A6: mov edx, [eax]
  loc_0060C3A8: push ecx
  loc_0060C3A9: push eax
  loc_0060C3AA: call [edx+0000008Ch]
  loc_0060C3B0: test eax, eax
  loc_0060C3B2: fnclex
  loc_0060C3B4: jge 0060C3CAh
  loc_0060C3B6: mov edx, var_C4
  loc_0060C3BC: push 0000008Ch
  loc_0060C3C1: push 0043B5D0h
  loc_0060C3C6: push edx
  loc_0060C3C7: push eax
  loc_0060C3C8: call edi
  loc_0060C3CA: lea ebx, var_28
  loc_0060C3CD: mov eax, var_24
  loc_0060C3D0: push ebx
  loc_0060C3D1: mov edx, 00000008h
  loc_0060C3D6: sub esp, 00000010h
  loc_0060C3D9: mov var_70, edx
  loc_0060C3DC: mov ebx, esp
  loc_0060C3DE: mov ecx, 0043B848h ; "Section4"
  loc_0060C3E3: mov var_68, ecx
  loc_0060C3E6: mov esi, [eax]
  loc_0060C3E8: mov [ebx], edx
  loc_0060C3EA: mov edx, var_6C
  loc_0060C3ED: push eax
  loc_0060C3EE: mov var_AC, eax
  loc_0060C3F4: mov [ebx+00000004h], edx
  loc_0060C3F7: mov [ebx+00000008h], ecx
  loc_0060C3FA: mov ecx, var_64
  loc_0060C3FD: mov [ebx+0000000Ch], ecx
  loc_0060C400: call [esi+0000001Ch]
  loc_0060C403: test eax, eax
  loc_0060C405: fnclex
  loc_0060C407: jge 0060C41Ah
  loc_0060C409: mov edx, var_AC
  loc_0060C40F: push 0000001Ch
  loc_0060C411: push 00439DD8h
  loc_0060C416: push edx
  loc_0060C417: push eax
  loc_0060C418: call edi
  loc_0060C41A: mov eax, var_28
  loc_0060C41D: lea edx, var_2C
  loc_0060C420: push edx
  loc_0060C421: push eax
  loc_0060C422: mov ecx, [eax]
  loc_0060C424: mov esi, eax
  loc_0060C426: call [ecx+0000003Ch]
  loc_0060C429: test eax, eax
  loc_0060C42B: fnclex
  loc_0060C42D: jge 0060C43Ah
  loc_0060C42F: push 0000003Ch
  loc_0060C431: push 00439BF8h
  loc_0060C436: push esi
  loc_0060C437: push eax
  loc_0060C438: call edi
  loc_0060C43A: lea ebx, var_30
  loc_0060C43D: mov eax, var_2C
  loc_0060C440: push ebx
  loc_0060C441: mov edx, 00000008h
  loc_0060C446: sub esp, 00000010h
  loc_0060C449: mov esi, [eax]
  loc_0060C44B: mov ebx, esp
  loc_0060C44D: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060C452: push eax
  loc_0060C453: mov var_BC, eax
  loc_0060C459: mov [ebx], edx
  loc_0060C45B: mov edx, var_7C
  loc_0060C45E: mov [ebx+00000004h], edx
  loc_0060C461: mov [ebx+00000008h], ecx
  loc_0060C464: mov ecx, var_74
  loc_0060C467: mov [ebx+0000000Ch], ecx
  loc_0060C46A: call [esi+0000001Ch]
  loc_0060C46D: test eax, eax
  loc_0060C46F: fnclex
  loc_0060C471: jge 0060C484h
  loc_0060C473: mov edx, var_BC
  loc_0060C479: push 0000001Ch
  loc_0060C47B: push 0043B7D8h
  loc_0060C480: push edx
  loc_0060C481: push eax
  loc_0060C482: call edi
  loc_0060C484: mov ecx, [006680F8h]
  loc_0060C48A: mov edx, [006680FCh]
  loc_0060C490: sub esp, 00000010h
  loc_0060C493: mov eax, esp
  loc_0060C495: push 0043B85Ch ; "Caption"
  loc_0060C49A: mov [eax], ecx
  loc_0060C49C: mov ecx, [00668100h]
  loc_0060C4A2: mov [eax+00000004h], edx
  loc_0060C4A5: mov edx, [00668104h]
  loc_0060C4AB: mov [eax+00000008h], ecx
  loc_0060C4AE: mov [eax+0000000Ch], edx
  loc_0060C4B1: mov eax, var_30
  loc_0060C4B4: push eax
  loc_0060C4B5: call [00401094h] ; __vbaLateMemSt
  loc_0060C4BB: lea ecx, var_30
  loc_0060C4BE: lea edx, var_2C
  loc_0060C4C1: push ecx
  loc_0060C4C2: lea eax, var_28
  loc_0060C4C5: push edx
  loc_0060C4C6: lea ecx, var_24
  loc_0060C4C9: push eax
  loc_0060C4CA: push ecx
  loc_0060C4CB: push 00000004h
  loc_0060C4CD: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C4D3: mov eax, var_C4
  loc_0060C4D9: add esp, 00000014h
  loc_0060C4DC: lea ecx, var_24
  loc_0060C4DF: mov edx, [eax]
  loc_0060C4E1: push ecx
  loc_0060C4E2: push eax
  loc_0060C4E3: call [edx+0000008Ch]
  loc_0060C4E9: test eax, eax
  loc_0060C4EB: fnclex
  loc_0060C4ED: jge 0060C503h
  loc_0060C4EF: mov edx, var_C4
  loc_0060C4F5: push 0000008Ch
  loc_0060C4FA: push 0043B5D0h
  loc_0060C4FF: push edx
  loc_0060C500: push eax
  loc_0060C501: call edi
  loc_0060C503: lea ebx, var_28
  loc_0060C506: mov eax, var_24
  loc_0060C509: push ebx
  loc_0060C50A: mov edx, 00000008h
  loc_0060C50F: sub esp, 00000010h
  loc_0060C512: mov var_70, edx
  loc_0060C515: mov ebx, esp
  loc_0060C517: mov ecx, 0043B848h ; "Section4"
  loc_0060C51C: mov var_68, ecx
  loc_0060C51F: mov esi, [eax]
  loc_0060C521: mov [ebx], edx
  loc_0060C523: mov edx, var_6C
  loc_0060C526: push eax
  loc_0060C527: mov var_AC, eax
  loc_0060C52D: mov [ebx+00000004h], edx
  loc_0060C530: mov [ebx+00000008h], ecx
  loc_0060C533: mov ecx, var_64
  loc_0060C536: mov [ebx+0000000Ch], ecx
  loc_0060C539: call [esi+0000001Ch]
  loc_0060C53C: test eax, eax
  loc_0060C53E: fnclex
  loc_0060C540: jge 0060C553h
  loc_0060C542: mov edx, var_AC
  loc_0060C548: push 0000001Ch
  loc_0060C54A: push 00439DD8h
  loc_0060C54F: push edx
  loc_0060C550: push eax
  loc_0060C551: call edi
  loc_0060C553: mov eax, var_28
  loc_0060C556: lea edx, var_2C
  loc_0060C559: push edx
  loc_0060C55A: push eax
  loc_0060C55B: mov ecx, [eax]
  loc_0060C55D: mov esi, eax
  loc_0060C55F: call [ecx+0000003Ch]
  loc_0060C562: test eax, eax
  loc_0060C564: fnclex
  loc_0060C566: jge 0060C573h
  loc_0060C568: push 0000003Ch
  loc_0060C56A: push 00439BF8h
  loc_0060C56F: push esi
  loc_0060C570: push eax
  loc_0060C571: call edi
  loc_0060C573: lea ebx, var_30
  loc_0060C576: mov eax, var_2C
  loc_0060C579: push ebx
  loc_0060C57A: mov edx, 00000008h
  loc_0060C57F: sub esp, 00000010h
  loc_0060C582: mov esi, [eax]
  loc_0060C584: mov ebx, esp
  loc_0060C586: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060C58B: push eax
  loc_0060C58C: mov var_BC, eax
  loc_0060C592: mov [ebx], edx
  loc_0060C594: mov edx, var_7C
  loc_0060C597: mov [ebx+00000004h], edx
  loc_0060C59A: mov [ebx+00000008h], ecx
  loc_0060C59D: mov ecx, var_74
  loc_0060C5A0: mov [ebx+0000000Ch], ecx
  loc_0060C5A3: call [esi+0000001Ch]
  loc_0060C5A6: test eax, eax
  loc_0060C5A8: fnclex
  loc_0060C5AA: jge 0060C5BDh
  loc_0060C5AC: mov edx, var_BC
  loc_0060C5B2: push 0000001Ch
  loc_0060C5B4: push 0043B7D8h
  loc_0060C5B9: push edx
  loc_0060C5BA: push eax
  loc_0060C5BB: call edi
  loc_0060C5BD: mov ecx, [00668108h]
  loc_0060C5C3: mov edx, [0066810Ch]
  loc_0060C5C9: sub esp, 00000010h
  loc_0060C5CC: mov eax, esp
  loc_0060C5CE: push 0043B85Ch ; "Caption"
  loc_0060C5D3: mov [eax], ecx
  loc_0060C5D5: mov ecx, [00668110h]
  loc_0060C5DB: mov [eax+00000004h], edx
  loc_0060C5DE: mov edx, [00668114h]
  loc_0060C5E4: mov [eax+00000008h], ecx
  loc_0060C5E7: mov [eax+0000000Ch], edx
  loc_0060C5EA: mov eax, var_30
  loc_0060C5ED: push eax
  loc_0060C5EE: call [00401094h] ; __vbaLateMemSt
  loc_0060C5F4: lea ecx, var_30
  loc_0060C5F7: lea edx, var_2C
  loc_0060C5FA: push ecx
  loc_0060C5FB: lea eax, var_28
  loc_0060C5FE: push edx
  loc_0060C5FF: lea ecx, var_24
  loc_0060C602: push eax
  loc_0060C603: push ecx
  loc_0060C604: push 00000004h
  loc_0060C606: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C60C: mov eax, var_C4
  loc_0060C612: add esp, 00000014h
  loc_0060C615: mov edx, [eax]
  loc_0060C617: push 0000000Ah
  loc_0060C619: push eax
  loc_0060C61A: call [edx+00000048h]
  loc_0060C61D: test eax, eax
  loc_0060C61F: fnclex
  loc_0060C621: jge 0060C634h
  loc_0060C623: mov ecx, var_C4
  loc_0060C629: push 00000048h
  loc_0060C62B: push 0043B5D0h
  loc_0060C630: push ecx
  loc_0060C631: push eax
  loc_0060C632: call edi
  loc_0060C634: mov eax, var_C4
  loc_0060C63A: push 0000000Ah
  loc_0060C63C: push eax
  loc_0060C63D: mov edx, [eax]
  loc_0060C63F: call [edx+00000058h]
  loc_0060C642: test eax, eax
  loc_0060C644: fnclex
  loc_0060C646: jge 0060C659h
  loc_0060C648: mov ecx, var_C4
  loc_0060C64E: push 00000058h
  loc_0060C650: push 0043B5D0h
  loc_0060C655: push ecx
  loc_0060C656: push eax
  loc_0060C657: call edi
  loc_0060C659: sub esp, 00000010h
  loc_0060C65C: mov eax, 00000002h
  loc_0060C661: mov edx, esp
  loc_0060C663: mov ecx, eax
  loc_0060C665: mov var_70, ecx
  loc_0060C668: mov var_68, eax
  loc_0060C66B: mov [edx], ecx
  loc_0060C66D: mov ecx, var_6C
  loc_0060C670: push 8001000Eh
  loc_0060C675: mov [edx+00000004h], ecx
  loc_0060C678: mov ecx, var_C4
  loc_0060C67E: push ecx
  loc_0060C67F: mov [edx+00000008h], eax
  loc_0060C682: mov eax, var_64
  loc_0060C685: mov [edx+0000000Ch], eax
  loc_0060C688: call [00401280h] ; __vbaLateIdSt
  loc_0060C68E: mov edx, var_C4
  loc_0060C694: push 00000000h
  loc_0060C696: push 80011003h
  loc_0060C69B: push edx
  loc_0060C69C: call [0040102Ch] ; __vbaLateIdCall
  loc_0060C6A2: add esp, 0000000Ch
  loc_0060C6A5: lea eax, var_C4
  loc_0060C6AB: push 00000000h
  loc_0060C6AD: push eax
  loc_0060C6AE: call [004010C8h] ; __vbaObjSetAddref
  loc_0060C6B4: push 0060C71Eh
  loc_0060C6B9: jmp 0060C6F4h
  loc_0060C6BB: lea ecx, var_20
  loc_0060C6BE: call [0040129Ch] ; __vbaFreeStr
  loc_0060C6C4: lea ecx, var_30
  loc_0060C6C7: lea edx, var_2C
  loc_0060C6CA: push ecx
  loc_0060C6CB: lea eax, var_28
  loc_0060C6CE: push edx
  loc_0060C6CF: lea ecx, var_24
  loc_0060C6D2: push eax
  loc_0060C6D3: push ecx
  loc_0060C6D4: push 00000004h
  loc_0060C6D6: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C6DC: lea edx, var_60
  loc_0060C6DF: lea eax, var_50
  loc_0060C6E2: push edx
  loc_0060C6E3: lea ecx, var_40
  loc_0060C6E6: push eax
  loc_0060C6E7: push ecx
  loc_0060C6E8: push 00000003h
  loc_0060C6EA: call [0040103Ch] ; __vbaFreeVarList
  loc_0060C6F0: add esp, 00000024h
  loc_0060C6F3: ret
  loc_0060C6F4: lea edx, var_C8
  loc_0060C6FA: lea eax, var_C4
  loc_0060C700: push edx
  loc_0060C701: push eax
  loc_0060C702: push 00000002h
  loc_0060C704: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C70A: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060C710: add esp, 0000000Ch
  loc_0060C713: lea ecx, var_18
  loc_0060C716: call __vbaFreeStr
  loc_0060C718: lea ecx, var_1C
  loc_0060C71B: call __vbaFreeStr
  loc_0060C71D: ret
  loc_0060C71E: mov ecx, var_10
  loc_0060C721: pop edi
  loc_0060C722: pop esi
  loc_0060C723: mov fs:[00000000h], ecx
  loc_0060C72A: pop ebx
  loc_0060C72B: mov esp, ebp
  loc_0060C72D: pop ebp
  loc_0060C72E: ret
End Sub

Private Sub Proc_34_3_60C740
  loc_0060C740: push ebp
  loc_0060C741: mov ebp, esp
  loc_0060C743: sub esp, 00000008h
  loc_0060C746: push 00405696h ; __vbaExceptHandler
  loc_0060C74B: mov eax, fs:[00000000h]
  loc_0060C751: push eax
  loc_0060C752: mov fs:[00000000h], esp
  loc_0060C759: sub esp, 000000D0h
  loc_0060C75F: push ebx
  loc_0060C760: push esi
  loc_0060C761: push edi
  loc_0060C762: mov var_8, esp
  loc_0060C765: mov var_4, 004044D0h
  loc_0060C76C: xor ebx, ebx
  loc_0060C76E: mov var_18, ebx
  loc_0060C771: mov var_1C, ebx
  loc_0060C774: mov var_20, ebx
  loc_0060C777: mov var_24, ebx
  loc_0060C77A: mov var_28, ebx
  loc_0060C77D: mov var_2C, ebx
  loc_0060C780: mov var_30, ebx
  loc_0060C783: mov var_40, ebx
  loc_0060C786: mov var_50, ebx
  loc_0060C789: mov var_60, ebx
  loc_0060C78C: mov var_70, ebx
  loc_0060C78F: mov var_80, ebx
  loc_0060C792: mov var_90, ebx
  loc_0060C798: mov var_A4, ebx
  loc_0060C79E: mov var_C4, ebx
  loc_0060C7A4: mov var_C8, ebx
  loc_0060C7AA: call 0055B320h
  loc_0060C7AF: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0060C7B5: mov edx, 0042DFECh
  loc_0060C7BA: lea ecx, var_18
  loc_0060C7BD: call __vbaStrCopy
  loc_0060C7BF: mov edx, 0043BA90h ; " SELECT kode_golongan, nama_golongan FROM brg_golongan ORDER BY kode_golongan"
  loc_0060C7C4: lea ecx, var_18
  loc_0060C7C7: call __vbaStrCopy
  loc_0060C7C9: push 0042DF30h
  loc_0060C7CE: call [00401168h] ; __vbaNew
  loc_0060C7D4: push eax
  loc_0060C7D5: push 00668090h
  loc_0060C7DA: call [004010B8h] ; __vbaObjSet
  loc_0060C7E0: cmp [0066803Ch], ebx
  loc_0060C7E6: jnz 0060C7F8h
  loc_0060C7E8: push 0066803Ch
  loc_0060C7ED: push 0042DEFCh
  loc_0060C7F2: call [004011E8h] ; __vbaNew2
  loc_0060C7F8: mov esi, [0066803Ch]
  loc_0060C7FE: lea ecx, var_40
  loc_0060C801: mov var_38, 80020004h
  loc_0060C808: mov var_40, 0000000Ah
  loc_0060C80F: call [0040123Ch] ; __vbaFreeVarg
  loc_0060C815: mov eax, [esi]
  loc_0060C817: lea ecx, var_24
  loc_0060C81A: push ecx
  loc_0060C81B: mov ecx, var_18
  loc_0060C81E: lea edx, var_40
  loc_0060C821: push 00000001h
  loc_0060C823: push edx
  loc_0060C824: push ecx
  loc_0060C825: push esi
  loc_0060C826: call [eax+00000040h]
  loc_0060C829: cmp eax, ebx
  loc_0060C82B: fnclex
  loc_0060C82D: jge 0060C842h
  loc_0060C82F: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060C835: push 00000040h
  loc_0060C837: push 0042E1B0h
  loc_0060C83C: push esi
  loc_0060C83D: push eax
  loc_0060C83E: call edi
  loc_0060C840: jmp 0060C848h
  loc_0060C842: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060C848: mov edx, var_24
  loc_0060C84B: mov esi, [00401268h] ; __vbaCastObj
  loc_0060C851: push 0042DF20h
  loc_0060C856: push edx
  loc_0060C857: call __vbaCastObj
  loc_0060C859: push eax
  loc_0060C85A: push 00668090h
  loc_0060C85F: call [004010B8h] ; __vbaObjSet
  loc_0060C865: lea ecx, var_24
  loc_0060C868: call [004012A0h] ; __vbaFreeObj
  loc_0060C86E: lea ecx, var_40
  loc_0060C871: call [00401020h] ; __vbaFreeVar
  loc_0060C877: cmp [006684F4h], ebx
  loc_0060C87D: jnz 0060C88Fh
  loc_0060C87F: push 006684F4h
  loc_0060C884: push 004083F4h
  loc_0060C889: call [004011E8h] ; __vbaNew2
  loc_0060C88F: mov eax, [006684F4h]
  loc_0060C894: lea ecx, var_C4
  loc_0060C89A: push eax
  loc_0060C89B: push ecx
  loc_0060C89C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060C8A2: mov edx, var_C4
  loc_0060C8A8: push 0043B7B8h
  loc_0060C8AD: push ebx
  loc_0060C8AE: mov edx, [edx]
  loc_0060C8B0: mov var_E0, edx
  loc_0060C8B6: call __vbaCastObj
  loc_0060C8B8: push eax
  loc_0060C8B9: lea eax, var_24
  loc_0060C8BC: push eax
  loc_0060C8BD: call [004010B8h] ; __vbaObjSet
  loc_0060C8C3: mov ecx, var_C4
  loc_0060C8C9: mov edx, var_E0
  loc_0060C8CF: push eax
  loc_0060C8D0: push ecx
  loc_0060C8D1: call [edx+00000028h]
  loc_0060C8D4: cmp eax, ebx
  loc_0060C8D6: fnclex
  loc_0060C8D8: jge 0060C8EBh
  loc_0060C8DA: mov ecx, var_C4
  loc_0060C8E0: push 00000028h
  loc_0060C8E2: push 0043B5D0h
  loc_0060C8E7: push ecx
  loc_0060C8E8: push eax
  loc_0060C8E9: call edi
  loc_0060C8EB: lea ecx, var_24
  loc_0060C8EE: call [004012A0h] ; __vbaFreeObj
  loc_0060C8F4: mov eax, var_C4
  loc_0060C8FA: push 0042DFECh
  loc_0060C8FF: push eax
  loc_0060C900: mov edx, [eax]
  loc_0060C902: call [edx+00000030h]
  loc_0060C905: cmp eax, ebx
  loc_0060C907: fnclex
  loc_0060C909: jge 0060C91Ch
  loc_0060C90B: mov ecx, var_C4
  loc_0060C911: push 00000030h
  loc_0060C913: push 0043B5D0h
  loc_0060C918: push ecx
  loc_0060C919: push eax
  loc_0060C91A: call edi
  loc_0060C91C: mov eax, [00668090h]
  loc_0060C921: lea ecx, var_24
  loc_0060C924: push ecx
  loc_0060C925: push eax
  loc_0060C926: mov edx, [eax]
  loc_0060C928: call [edx+00000114h]
  loc_0060C92E: cmp eax, ebx
  loc_0060C930: fnclex
  loc_0060C932: jge 0060C948h
  loc_0060C934: mov edx, [00668090h]
  loc_0060C93A: push 00000114h
  loc_0060C93F: push 0043B7C8h
  loc_0060C944: push edx
  loc_0060C945: push eax
  loc_0060C946: call edi
  loc_0060C948: mov eax, var_C4
  loc_0060C94E: mov ecx, var_24
  loc_0060C951: push 0043B7B8h
  loc_0060C956: push ecx
  loc_0060C957: mov esi, [eax]
  loc_0060C959: call [00401268h] ; __vbaCastObj
  loc_0060C95F: lea edx, var_28
  loc_0060C962: push eax
  loc_0060C963: push edx
  loc_0060C964: call [004010B8h] ; __vbaObjSet
  loc_0060C96A: push eax
  loc_0060C96B: mov eax, var_C4
  loc_0060C971: push eax
  loc_0060C972: call [esi+00000028h]
  loc_0060C975: cmp eax, ebx
  loc_0060C977: fnclex
  loc_0060C979: jge 0060C98Ch
  loc_0060C97B: mov ecx, var_C4
  loc_0060C981: push 00000028h
  loc_0060C983: push 0043B5D0h
  loc_0060C988: push ecx
  loc_0060C989: push eax
  loc_0060C98A: call edi
  loc_0060C98C: lea edx, var_28
  loc_0060C98F: lea eax, var_24
  loc_0060C992: push edx
  loc_0060C993: push eax
  loc_0060C994: push 00000002h
  loc_0060C996: call [0040104Ch] ; __vbaFreeObjList
  loc_0060C99C: mov eax, var_C4
  loc_0060C9A2: add esp, 0000000Ch
  loc_0060C9A5: lea edx, var_24
  loc_0060C9A8: mov ecx, [eax]
  loc_0060C9AA: push edx
  loc_0060C9AB: push eax
  loc_0060C9AC: call [ecx+0000008Ch]
  loc_0060C9B2: cmp eax, ebx
  loc_0060C9B4: fnclex
  loc_0060C9B6: jge 0060C9CCh
  loc_0060C9B8: mov ecx, var_C4
  loc_0060C9BE: push 0000008Ch
  loc_0060C9C3: push 0043B5D0h
  loc_0060C9C8: push ecx
  loc_0060C9C9: push eax
  loc_0060C9CA: call edi
  loc_0060C9CC: lea esi, var_28
  loc_0060C9CF: mov eax, var_24
  loc_0060C9D2: push esi
  loc_0060C9D3: mov ecx, 00000008h
  loc_0060C9D8: sub esp, 00000010h
  loc_0060C9DB: mov var_70, ecx
  loc_0060C9DE: mov esi, esp
  loc_0060C9E0: mov var_68, 00439FB0h ; "Section1"
  loc_0060C9E7: mov edx, [eax]
  loc_0060C9E9: push eax
  loc_0060C9EA: mov [esi], ecx
  loc_0060C9EC: mov ecx, var_6C
  loc_0060C9EF: mov var_AC, eax
  loc_0060C9F5: mov [esi+00000004h], ecx
  loc_0060C9F8: mov ecx, var_68
  loc_0060C9FB: mov [esi+00000008h], ecx
  loc_0060C9FE: mov ecx, var_64
  loc_0060CA01: mov [esi+0000000Ch], ecx
  loc_0060CA04: call [edx+0000001Ch]
  loc_0060CA07: cmp eax, ebx
  loc_0060CA09: fnclex
  loc_0060CA0B: jge 0060CA1Eh
  loc_0060CA0D: mov edx, var_AC
  loc_0060CA13: push 0000001Ch
  loc_0060CA15: push 00439DD8h
  loc_0060CA1A: push edx
  loc_0060CA1B: push eax
  loc_0060CA1C: call edi
  loc_0060CA1E: mov eax, var_28
  loc_0060CA21: lea edx, var_2C
  loc_0060CA24: push edx
  loc_0060CA25: push eax
  loc_0060CA26: mov ecx, [eax]
  loc_0060CA28: mov esi, eax
  loc_0060CA2A: call [ecx+0000003Ch]
  loc_0060CA2D: cmp eax, ebx
  loc_0060CA2F: fnclex
  loc_0060CA31: jge 0060CA3Eh
  loc_0060CA33: push 0000003Ch
  loc_0060CA35: push 00439BF8h
  loc_0060CA3A: push esi
  loc_0060CA3B: push eax
  loc_0060CA3C: call edi
  loc_0060CA3E: mov eax, var_2C
  loc_0060CA41: mov var_2C, ebx
  loc_0060CA44: push eax
  loc_0060CA45: lea eax, var_C8
  loc_0060CA4B: push eax
  loc_0060CA4C: call [004010B8h] ; __vbaObjSet
  loc_0060CA52: lea ecx, var_28
  loc_0060CA55: lea edx, var_24
  loc_0060CA58: push ecx
  loc_0060CA59: push edx
  loc_0060CA5A: push 00000002h
  loc_0060CA5C: call [0040104Ch] ; __vbaFreeObjList
  loc_0060CA62: mov eax, var_C8
  loc_0060CA68: add esp, 0000000Ch
  loc_0060CA6B: lea edx, var_A4
  loc_0060CA71: mov ecx, [eax]
  loc_0060CA73: push edx
  loc_0060CA74: push eax
  loc_0060CA75: call [ecx+00000020h]
  loc_0060CA78: cmp eax, ebx
  loc_0060CA7A: fnclex
  loc_0060CA7C: jge 0060CA8Fh
  loc_0060CA7E: mov ecx, var_C8
  loc_0060CA84: push 00000020h
  loc_0060CA86: push 0043B7D8h
  loc_0060CA8B: push ecx
  loc_0060CA8C: push eax
  loc_0060CA8D: call edi
  loc_0060CA8F: mov ecx, var_A4
  loc_0060CA95: call [00401138h] ; __vbaI2I4
  loc_0060CA9B: mov var_D0, eax
  loc_0060CAA1: mov ebx, 00000001h
  loc_0060CAA6: cmp bx, var_D0
  loc_0060CAAD: mov var_14, ebx
  loc_0060CAB0: jg 0060CD21h
  loc_0060CAB6: lea esi, var_24
  loc_0060CAB9: mov ecx, var_C8
  loc_0060CABF: push esi
  loc_0060CAC0: mov eax, 00000002h
  loc_0060CAC5: sub esp, 00000010h
  loc_0060CAC8: mov var_70, eax
  loc_0060CACB: mov esi, esp
  loc_0060CACD: mov var_68, bx
  loc_0060CAD1: mov edx, [ecx]
  loc_0060CAD3: push ecx
  loc_0060CAD4: mov [esi], eax
  loc_0060CAD6: mov eax, var_6C
  loc_0060CAD9: mov [esi+00000004h], eax
  loc_0060CADC: mov eax, var_68
  loc_0060CADF: mov [esi+00000008h], eax
  loc_0060CAE2: mov eax, var_64
  loc_0060CAE5: mov [esi+0000000Ch], eax
  loc_0060CAE8: call [edx+0000001Ch]
  loc_0060CAEB: test eax, eax
  loc_0060CAED: fnclex
  loc_0060CAEF: jge 0060CB02h
  loc_0060CAF1: mov ecx, var_C8
  loc_0060CAF7: push 0000001Ch
  loc_0060CAF9: push 0043B7D8h
  loc_0060CAFE: push ecx
  loc_0060CAFF: push eax
  loc_0060CB00: call edi
  loc_0060CB02: mov edx, var_24
  loc_0060CB05: push 0043B7E8h
  loc_0060CB0A: push edx
  loc_0060CB0B: call [004011CCh] ; __vbaCheckType
  loc_0060CB11: lea ecx, var_24
  loc_0060CB14: mov si, ax
  loc_0060CB17: call [004012A0h] ; __vbaFreeObj
  loc_0060CB1D: test si, si
  loc_0060CB20: Unknown_795()
  loc_0060CB26: mov var_68, bx
  loc_0060CB2A: lea ebx, var_24
  loc_0060CB2D: push ebx
  loc_0060CB2E: mov ecx, var_C8
  loc_0060CB34: sub esp, 00000010h
  loc_0060CB37: mov eax, 00000002h
  loc_0060CB3C: mov ebx, esp
  loc_0060CB3E: mov var_70, eax
  loc_0060CB41: mov edx, [ecx]
  loc_0060CB43: mov esi, 0042DFECh
  loc_0060CB48: mov [ebx], eax
  loc_0060CB4A: mov eax, var_6C
  loc_0060CB4D: push ecx
  loc_0060CB4E: mov var_78, esi
  loc_0060CB51: mov [ebx+00000004h], eax
  loc_0060CB54: mov eax, var_68
  loc_0060CB57: mov edi, 00000008h
  loc_0060CB5C: mov [ebx+00000008h], eax
  loc_0060CB5F: mov eax, var_64
  loc_0060CB62: mov [ebx+0000000Ch], eax
  loc_0060CB65: call [edx+0000001Ch]
  loc_0060CB68: test eax, eax
  loc_0060CB6A: fnclex
  loc_0060CB6C: jge 0060CB83h
  loc_0060CB6E: mov ecx, var_C8
  loc_0060CB74: push 0000001Ch
  loc_0060CB76: push 0043B7D8h
  loc_0060CB7B: push ecx
  loc_0060CB7C: push eax
  loc_0060CB7D: call [00401080h] ; __vbaHresultCheckObj
  loc_0060CB83: mov eax, var_7C
  loc_0060CB86: sub esp, 00000010h
  loc_0060CB89: mov edx, esp
  loc_0060CB8B: mov ecx, var_74
  loc_0060CB8E: push 0043B7F8h ; "DataMember"
  loc_0060CB93: mov [edx], edi
  loc_0060CB95: mov [edx+00000004h], eax
  loc_0060CB98: mov [edx+00000008h], esi
  loc_0060CB9B: mov [edx+0000000Ch], ecx
  loc_0060CB9E: mov edx, var_24
  loc_0060CBA1: push edx
  loc_0060CBA2: call [00401094h] ; __vbaLateMemSt
  loc_0060CBA8: lea ecx, var_24
  loc_0060CBAB: call [004012A0h] ; __vbaFreeObj
  loc_0060CBB1: mov eax, [00668090h]
  loc_0060CBB6: lea edx, var_24
  loc_0060CBB9: push edx
  loc_0060CBBA: push eax
  loc_0060CBBB: mov ecx, [eax]
  loc_0060CBBD: call [ecx+00000054h]
  loc_0060CBC0: test eax, eax
  loc_0060CBC2: fnclex
  loc_0060CBC4: jge 0060CBDBh
  loc_0060CBC6: mov ecx, [00668090h]
  loc_0060CBCC: push 00000054h
  loc_0060CBCE: push 0042DF44h
  loc_0060CBD3: push ecx
  loc_0060CBD4: push eax
  loc_0060CBD5: call [00401080h] ; __vbaHresultCheckObj
  loc_0060CBDB: mov ebx, var_14
  loc_0060CBDE: lea edi, var_28
  loc_0060CBE1: mov dx, bx
  loc_0060CBE4: push edi
  loc_0060CBE5: sub dx, 0001h
  loc_0060CBE9: mov eax, var_24
  loc_0060CBEC: jo 0060D4AFh
  loc_0060CBF2: sub esp, 00000010h
  loc_0060CBF5: mov ecx, 00000002h
  loc_0060CBFA: mov edi, esp
  loc_0060CBFC: mov var_70, ecx
  loc_0060CBFF: mov var_68, dx
  loc_0060CC03: mov edx, [eax]
  loc_0060CC05: mov [edi], ecx
  loc_0060CC07: mov ecx, var_6C
  loc_0060CC0A: push eax
  loc_0060CC0B: mov esi, eax
  loc_0060CC0D: mov [edi+00000004h], ecx
  loc_0060CC10: mov ecx, var_68
  loc_0060CC13: mov [edi+00000008h], ecx
  loc_0060CC16: mov ecx, var_64
  loc_0060CC19: mov [edi+0000000Ch], ecx
  loc_0060CC1C: call [edx+00000028h]
  loc_0060CC1F: test eax, eax
  loc_0060CC21: fnclex
  loc_0060CC23: jge 0060CC38h
  loc_0060CC25: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060CC2B: push 00000028h
  loc_0060CC2D: push 0042DFACh
  loc_0060CC32: push esi
  loc_0060CC33: push eax
  loc_0060CC34: call edi
  loc_0060CC36: jmp 0060CC3Eh
  loc_0060CC38: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060CC3E: mov eax, var_28
  loc_0060CC41: lea ecx, var_20
  loc_0060CC44: push ecx
  loc_0060CC45: push eax
  loc_0060CC46: mov edx, [eax]
  loc_0060CC48: mov esi, eax
  loc_0060CC4A: call [edx+0000002Ch]
  loc_0060CC4D: test eax, eax
  loc_0060CC4F: fnclex
  loc_0060CC51: jge 0060CC5Eh
  loc_0060CC53: push 0000002Ch
  loc_0060CC55: push 0042DFBCh
  loc_0060CC5A: push esi
  loc_0060CC5B: push eax
  loc_0060CC5C: call edi
  loc_0060CC5E: mov eax, var_20
  loc_0060CC61: lea esi, var_2C
  loc_0060CC64: push esi
  loc_0060CC65: mov ecx, var_C8
  loc_0060CC6B: sub esp, 00000010h
  loc_0060CC6E: mov var_38, eax
  loc_0060CC71: mov esi, esp
  loc_0060CC73: mov eax, 00000002h
  loc_0060CC78: mov var_78, bx
  loc_0060CC7C: mov var_20, 00000000h
  loc_0060CC83: mov [esi], eax
  loc_0060CC85: mov eax, var_7C
  loc_0060CC88: mov var_40, 00000008h
  loc_0060CC8F: mov edx, [ecx]
  loc_0060CC91: mov [esi+00000004h], eax
  loc_0060CC94: mov eax, var_78
  loc_0060CC97: push ecx
  loc_0060CC98: mov [esi+00000008h], eax
  loc_0060CC9B: mov eax, var_74
  loc_0060CC9E: mov [esi+0000000Ch], eax
  loc_0060CCA1: call [edx+0000001Ch]
  loc_0060CCA4: test eax, eax
  loc_0060CCA6: fnclex
  loc_0060CCA8: jge 0060CCBBh
  loc_0060CCAA: mov ecx, var_C8
  loc_0060CCB0: push 0000001Ch
  loc_0060CCB2: push 0043B7D8h
  loc_0060CCB7: push ecx
  loc_0060CCB8: push eax
  loc_0060CCB9: call edi
  loc_0060CCBB: mov eax, var_40
  loc_0060CCBE: mov ecx, var_3C
  loc_0060CCC1: sub esp, 00000010h
  loc_0060CCC4: mov edx, esp
  loc_0060CCC6: push 0043B810h ; "DataField"
  loc_0060CCCB: mov [edx], eax
  loc_0060CCCD: mov eax, var_38
  loc_0060CCD0: mov [edx+00000004h], ecx
  loc_0060CCD3: mov ecx, var_34
  loc_0060CCD6: mov [edx+00000008h], eax
  loc_0060CCD9: mov [edx+0000000Ch], ecx
  loc_0060CCDC: mov edx, var_2C
  loc_0060CCDF: push edx
  loc_0060CCE0: call [00401094h] ; __vbaLateMemSt
  loc_0060CCE6: lea eax, var_2C
  loc_0060CCE9: lea ecx, var_28
  loc_0060CCEC: push eax
  loc_0060CCED: lea edx, var_24
  loc_0060CCF0: push ecx
  loc_0060CCF1: push edx
  loc_0060CCF2: push 00000003h
  loc_0060CCF4: call [0040104Ch] ; __vbaFreeObjList
  loc_0060CCFA: add esp, 00000010h
  loc_0060CCFD: lea ecx, var_40
  loc_0060CD00: call [00401020h] ; __vbaFreeVar
  loc_0060CD06: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060CD0C: mov eax, 00000001h
  loc_0060CD11: add ax, bx
  loc_0060CD14: jo 0060D4AFh
  loc_0060CD1A: mov ebx, eax
  loc_0060CD1C: jmp 0060CAA6h
  loc_0060CD21: lea eax, var_C8
  loc_0060CD27: push 00000000h
  loc_0060CD29: push eax
  loc_0060CD2A: call [004010C8h] ; __vbaObjSetAddref
  loc_0060CD30: lea ecx, var_40
  loc_0060CD33: push ecx
  loc_0060CD34: call [00401220h] ; rtcGetDateVar
  loc_0060CD3A: lea edx, var_70
  loc_0060CD3D: lea ecx, var_50
  loc_0060CD40: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060CD47: mov var_70, 00000008h
  loc_0060CD4E: call [00401238h] ; __vbaVarDup
  loc_0060CD54: push 00000001h
  loc_0060CD56: lea edx, var_50
  loc_0060CD59: push 00000001h
  loc_0060CD5B: lea eax, var_40
  loc_0060CD5E: push edx
  loc_0060CD5F: lea ecx, var_60
  loc_0060CD62: push eax
  loc_0060CD63: push ecx
  loc_0060CD64: call [00401078h] ; rtcVarFromFormatVar
  loc_0060CD6A: mov eax, var_C4
  loc_0060CD70: lea ecx, var_24
  loc_0060CD73: push ecx
  loc_0060CD74: push eax
  loc_0060CD75: mov edx, [eax]
  loc_0060CD77: call [edx+0000008Ch]
  loc_0060CD7D: test eax, eax
  loc_0060CD7F: fnclex
  loc_0060CD81: jge 0060CD97h
  loc_0060CD83: mov edx, var_C4
  loc_0060CD89: push 0000008Ch
  loc_0060CD8E: push 0043B5D0h
  loc_0060CD93: push edx
  loc_0060CD94: push eax
  loc_0060CD95: call edi
  loc_0060CD97: lea ebx, var_28
  loc_0060CD9A: mov eax, var_24
  loc_0060CD9D: push ebx
  loc_0060CD9E: mov edx, 00000008h
  loc_0060CDA3: sub esp, 00000010h
  loc_0060CDA6: mov esi, [eax]
  loc_0060CDA8: mov ebx, esp
  loc_0060CDAA: mov ecx, 0043B848h ; "Section4"
  loc_0060CDAF: push eax
  loc_0060CDB0: mov var_AC, eax
  loc_0060CDB6: mov [ebx], edx
  loc_0060CDB8: mov edx, var_7C
  loc_0060CDBB: mov [ebx+00000004h], edx
  loc_0060CDBE: mov [ebx+00000008h], ecx
  loc_0060CDC1: mov ecx, var_74
  loc_0060CDC4: mov [ebx+0000000Ch], ecx
  loc_0060CDC7: call [esi+0000001Ch]
  loc_0060CDCA: test eax, eax
  loc_0060CDCC: fnclex
  loc_0060CDCE: jge 0060CDE1h
  loc_0060CDD0: mov edx, var_AC
  loc_0060CDD6: push 0000001Ch
  loc_0060CDD8: push 00439DD8h
  loc_0060CDDD: push edx
  loc_0060CDDE: push eax
  loc_0060CDDF: call edi
  loc_0060CDE1: mov eax, var_28
  loc_0060CDE4: lea edx, var_2C
  loc_0060CDE7: push edx
  loc_0060CDE8: push eax
  loc_0060CDE9: mov ecx, [eax]
  loc_0060CDEB: mov esi, eax
  loc_0060CDED: call [ecx+0000003Ch]
  loc_0060CDF0: test eax, eax
  loc_0060CDF2: fnclex
  loc_0060CDF4: jge 0060CE01h
  loc_0060CDF6: push 0000003Ch
  loc_0060CDF8: push 00439BF8h
  loc_0060CDFD: push esi
  loc_0060CDFE: push eax
  loc_0060CDFF: call edi
  loc_0060CE01: lea ebx, var_30
  loc_0060CE04: mov eax, var_2C
  loc_0060CE07: push ebx
  loc_0060CE08: mov edx, 00000008h
  loc_0060CE0D: sub esp, 00000010h
  loc_0060CE10: mov esi, [eax]
  loc_0060CE12: mov ebx, esp
  loc_0060CE14: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060CE19: push eax
  loc_0060CE1A: mov var_BC, eax
  loc_0060CE20: mov [ebx], edx
  loc_0060CE22: mov edx, var_8C
  loc_0060CE28: mov [ebx+00000004h], edx
  loc_0060CE2B: mov [ebx+00000008h], ecx
  loc_0060CE2E: mov ecx, var_84
  loc_0060CE34: mov [ebx+0000000Ch], ecx
  loc_0060CE37: call [esi+0000001Ch]
  loc_0060CE3A: test eax, eax
  loc_0060CE3C: fnclex
  loc_0060CE3E: jge 0060CE51h
  loc_0060CE40: mov edx, var_BC
  loc_0060CE46: push 0000001Ch
  loc_0060CE48: push 0043B7D8h
  loc_0060CE4D: push edx
  loc_0060CE4E: push eax
  loc_0060CE4F: call edi
  loc_0060CE51: mov ecx, var_60
  loc_0060CE54: mov edx, var_5C
  loc_0060CE57: sub esp, 00000010h
  loc_0060CE5A: mov eax, esp
  loc_0060CE5C: push 0043B85Ch ; "Caption"
  loc_0060CE61: mov [eax], ecx
  loc_0060CE63: mov ecx, var_58
  loc_0060CE66: mov [eax+00000004h], edx
  loc_0060CE69: mov edx, var_54
  loc_0060CE6C: mov [eax+00000008h], ecx
  loc_0060CE6F: mov [eax+0000000Ch], edx
  loc_0060CE72: mov eax, var_30
  loc_0060CE75: push eax
  loc_0060CE76: call [00401094h] ; __vbaLateMemSt
  loc_0060CE7C: lea ecx, var_30
  loc_0060CE7F: lea edx, var_2C
  loc_0060CE82: push ecx
  loc_0060CE83: lea eax, var_28
  loc_0060CE86: push edx
  loc_0060CE87: lea ecx, var_24
  loc_0060CE8A: push eax
  loc_0060CE8B: push ecx
  loc_0060CE8C: push 00000004h
  loc_0060CE8E: call [0040104Ch] ; __vbaFreeObjList
  loc_0060CE94: lea edx, var_60
  loc_0060CE97: lea eax, var_50
  loc_0060CE9A: push edx
  loc_0060CE9B: lea ecx, var_40
  loc_0060CE9E: push eax
  loc_0060CE9F: push ecx
  loc_0060CEA0: push 00000003h
  loc_0060CEA2: call [0040103Ch] ; __vbaFreeVarList
  loc_0060CEA8: mov eax, var_C4
  loc_0060CEAE: add esp, 00000024h
  loc_0060CEB1: lea ecx, var_24
  loc_0060CEB4: mov edx, [eax]
  loc_0060CEB6: push ecx
  loc_0060CEB7: push eax
  loc_0060CEB8: call [edx+0000008Ch]
  loc_0060CEBE: test eax, eax
  loc_0060CEC0: fnclex
  loc_0060CEC2: jge 0060CED8h
  loc_0060CEC4: mov edx, var_C4
  loc_0060CECA: push 0000008Ch
  loc_0060CECF: push 0043B5D0h
  loc_0060CED4: push edx
  loc_0060CED5: push eax
  loc_0060CED6: call edi
  loc_0060CED8: lea ebx, var_28
  loc_0060CEDB: mov eax, var_24
  loc_0060CEDE: push ebx
  loc_0060CEDF: mov edx, 00000008h
  loc_0060CEE4: sub esp, 00000010h
  loc_0060CEE7: mov var_70, edx
  loc_0060CEEA: mov ebx, esp
  loc_0060CEEC: mov ecx, 0043B848h ; "Section4"
  loc_0060CEF1: mov var_68, ecx
  loc_0060CEF4: mov esi, [eax]
  loc_0060CEF6: mov [ebx], edx
  loc_0060CEF8: mov edx, var_6C
  loc_0060CEFB: push eax
  loc_0060CEFC: mov var_AC, eax
  loc_0060CF02: mov [ebx+00000004h], edx
  loc_0060CF05: mov [ebx+00000008h], ecx
  loc_0060CF08: mov ecx, var_64
  loc_0060CF0B: mov [ebx+0000000Ch], ecx
  loc_0060CF0E: call [esi+0000001Ch]
  loc_0060CF11: test eax, eax
  loc_0060CF13: fnclex
  loc_0060CF15: jge 0060CF28h
  loc_0060CF17: mov edx, var_AC
  loc_0060CF1D: push 0000001Ch
  loc_0060CF1F: push 00439DD8h
  loc_0060CF24: push edx
  loc_0060CF25: push eax
  loc_0060CF26: call edi
  loc_0060CF28: mov eax, var_28
  loc_0060CF2B: lea edx, var_2C
  loc_0060CF2E: push edx
  loc_0060CF2F: push eax
  loc_0060CF30: mov ecx, [eax]
  loc_0060CF32: mov esi, eax
  loc_0060CF34: call [ecx+0000003Ch]
  loc_0060CF37: test eax, eax
  loc_0060CF39: fnclex
  loc_0060CF3B: jge 0060CF48h
  loc_0060CF3D: push 0000003Ch
  loc_0060CF3F: push 00439BF8h
  loc_0060CF44: push esi
  loc_0060CF45: push eax
  loc_0060CF46: call edi
  loc_0060CF48: lea ebx, var_30
  loc_0060CF4B: mov eax, var_2C
  loc_0060CF4E: push ebx
  loc_0060CF4F: mov edx, 00000008h
  loc_0060CF54: sub esp, 00000010h
  loc_0060CF57: mov esi, [eax]
  loc_0060CF59: mov ebx, esp
  loc_0060CF5B: mov ecx, 0043B870h ; "LabelNama"
  loc_0060CF60: push eax
  loc_0060CF61: mov var_BC, eax
  loc_0060CF67: mov [ebx], edx
  loc_0060CF69: mov edx, var_7C
  loc_0060CF6C: mov [ebx+00000004h], edx
  loc_0060CF6F: mov [ebx+00000008h], ecx
  loc_0060CF72: mov ecx, var_74
  loc_0060CF75: mov [ebx+0000000Ch], ecx
  loc_0060CF78: call [esi+0000001Ch]
  loc_0060CF7B: test eax, eax
  loc_0060CF7D: fnclex
  loc_0060CF7F: jge 0060CF92h
  loc_0060CF81: mov edx, var_BC
  loc_0060CF87: push 0000001Ch
  loc_0060CF89: push 0043B7D8h
  loc_0060CF8E: push edx
  loc_0060CF8F: push eax
  loc_0060CF90: call edi
  loc_0060CF92: mov ecx, [006680D8h]
  loc_0060CF98: mov edx, [006680DCh]
  loc_0060CF9E: sub esp, 00000010h
  loc_0060CFA1: mov eax, esp
  loc_0060CFA3: push 0043B85Ch ; "Caption"
  loc_0060CFA8: mov [eax], ecx
  loc_0060CFAA: mov ecx, [006680E0h]
  loc_0060CFB0: mov [eax+00000004h], edx
  loc_0060CFB3: mov edx, [006680E4h]
  loc_0060CFB9: mov [eax+00000008h], ecx
  loc_0060CFBC: mov [eax+0000000Ch], edx
  loc_0060CFBF: mov eax, var_30
  loc_0060CFC2: push eax
  loc_0060CFC3: call [00401094h] ; __vbaLateMemSt
  loc_0060CFC9: lea ecx, var_30
  loc_0060CFCC: lea edx, var_2C
  loc_0060CFCF: push ecx
  loc_0060CFD0: lea eax, var_28
  loc_0060CFD3: push edx
  loc_0060CFD4: lea ecx, var_24
  loc_0060CFD7: push eax
  loc_0060CFD8: push ecx
  loc_0060CFD9: push 00000004h
  loc_0060CFDB: call [0040104Ch] ; __vbaFreeObjList
  loc_0060CFE1: mov eax, var_C4
  loc_0060CFE7: add esp, 00000014h
  loc_0060CFEA: lea ecx, var_24
  loc_0060CFED: mov edx, [eax]
  loc_0060CFEF: push ecx
  loc_0060CFF0: push eax
  loc_0060CFF1: call [edx+0000008Ch]
  loc_0060CFF7: test eax, eax
  loc_0060CFF9: fnclex
  loc_0060CFFB: jge 0060D011h
  loc_0060CFFD: mov edx, var_C4
  loc_0060D003: push 0000008Ch
  loc_0060D008: push 0043B5D0h
  loc_0060D00D: push edx
  loc_0060D00E: push eax
  loc_0060D00F: call edi
  loc_0060D011: lea ebx, var_28
  loc_0060D014: mov eax, var_24
  loc_0060D017: push ebx
  loc_0060D018: mov edx, 00000008h
  loc_0060D01D: sub esp, 00000010h
  loc_0060D020: mov var_70, edx
  loc_0060D023: mov ebx, esp
  loc_0060D025: mov ecx, 0043B848h ; "Section4"
  loc_0060D02A: mov var_68, ecx
  loc_0060D02D: mov esi, [eax]
  loc_0060D02F: mov [ebx], edx
  loc_0060D031: mov edx, var_6C
  loc_0060D034: push eax
  loc_0060D035: mov var_AC, eax
  loc_0060D03B: mov [ebx+00000004h], edx
  loc_0060D03E: mov [ebx+00000008h], ecx
  loc_0060D041: mov ecx, var_64
  loc_0060D044: mov [ebx+0000000Ch], ecx
  loc_0060D047: call [esi+0000001Ch]
  loc_0060D04A: test eax, eax
  loc_0060D04C: fnclex
  loc_0060D04E: jge 0060D061h
  loc_0060D050: mov edx, var_AC
  loc_0060D056: push 0000001Ch
  loc_0060D058: push 00439DD8h
  loc_0060D05D: push edx
  loc_0060D05E: push eax
  loc_0060D05F: call edi
  loc_0060D061: mov eax, var_28
  loc_0060D064: lea edx, var_2C
  loc_0060D067: push edx
  loc_0060D068: push eax
  loc_0060D069: mov ecx, [eax]
  loc_0060D06B: mov esi, eax
  loc_0060D06D: call [ecx+0000003Ch]
  loc_0060D070: test eax, eax
  loc_0060D072: fnclex
  loc_0060D074: jge 0060D081h
  loc_0060D076: push 0000003Ch
  loc_0060D078: push 00439BF8h
  loc_0060D07D: push esi
  loc_0060D07E: push eax
  loc_0060D07F: call edi
  loc_0060D081: lea ebx, var_30
  loc_0060D084: mov eax, var_2C
  loc_0060D087: push ebx
  loc_0060D088: mov edx, 00000008h
  loc_0060D08D: sub esp, 00000010h
  loc_0060D090: mov esi, [eax]
  loc_0060D092: mov ebx, esp
  loc_0060D094: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060D099: push eax
  loc_0060D09A: mov var_BC, eax
  loc_0060D0A0: mov [ebx], edx
  loc_0060D0A2: mov edx, var_7C
  loc_0060D0A5: mov [ebx+00000004h], edx
  loc_0060D0A8: mov [ebx+00000008h], ecx
  loc_0060D0AB: mov ecx, var_74
  loc_0060D0AE: mov [ebx+0000000Ch], ecx
  loc_0060D0B1: call [esi+0000001Ch]
  loc_0060D0B4: test eax, eax
  loc_0060D0B6: fnclex
  loc_0060D0B8: jge 0060D0CBh
  loc_0060D0BA: mov edx, var_BC
  loc_0060D0C0: push 0000001Ch
  loc_0060D0C2: push 0043B7D8h
  loc_0060D0C7: push edx
  loc_0060D0C8: push eax
  loc_0060D0C9: call edi
  loc_0060D0CB: mov ecx, [006680E8h]
  loc_0060D0D1: mov edx, [006680ECh]
  loc_0060D0D7: sub esp, 00000010h
  loc_0060D0DA: mov eax, esp
  loc_0060D0DC: push 0043B85Ch ; "Caption"
  loc_0060D0E1: mov [eax], ecx
  loc_0060D0E3: mov ecx, [006680F0h]
  loc_0060D0E9: mov [eax+00000004h], edx
  loc_0060D0EC: mov edx, [006680F4h]
  loc_0060D0F2: mov [eax+00000008h], ecx
  loc_0060D0F5: mov [eax+0000000Ch], edx
  loc_0060D0F8: mov eax, var_30
  loc_0060D0FB: push eax
  loc_0060D0FC: call [00401094h] ; __vbaLateMemSt
  loc_0060D102: lea ecx, var_30
  loc_0060D105: lea edx, var_2C
  loc_0060D108: push ecx
  loc_0060D109: lea eax, var_28
  loc_0060D10C: push edx
  loc_0060D10D: lea ecx, var_24
  loc_0060D110: push eax
  loc_0060D111: push ecx
  loc_0060D112: push 00000004h
  loc_0060D114: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D11A: mov eax, var_C4
  loc_0060D120: add esp, 00000014h
  loc_0060D123: lea ecx, var_24
  loc_0060D126: mov edx, [eax]
  loc_0060D128: push ecx
  loc_0060D129: push eax
  loc_0060D12A: call [edx+0000008Ch]
  loc_0060D130: test eax, eax
  loc_0060D132: fnclex
  loc_0060D134: jge 0060D14Ah
  loc_0060D136: mov edx, var_C4
  loc_0060D13C: push 0000008Ch
  loc_0060D141: push 0043B5D0h
  loc_0060D146: push edx
  loc_0060D147: push eax
  loc_0060D148: call edi
  loc_0060D14A: lea ebx, var_28
  loc_0060D14D: mov eax, var_24
  loc_0060D150: push ebx
  loc_0060D151: mov edx, 00000008h
  loc_0060D156: sub esp, 00000010h
  loc_0060D159: mov var_70, edx
  loc_0060D15C: mov ebx, esp
  loc_0060D15E: mov ecx, 0043B848h ; "Section4"
  loc_0060D163: mov var_68, ecx
  loc_0060D166: mov esi, [eax]
  loc_0060D168: mov [ebx], edx
  loc_0060D16A: mov edx, var_6C
  loc_0060D16D: push eax
  loc_0060D16E: mov var_AC, eax
  loc_0060D174: mov [ebx+00000004h], edx
  loc_0060D177: mov [ebx+00000008h], ecx
  loc_0060D17A: mov ecx, var_64
  loc_0060D17D: mov [ebx+0000000Ch], ecx
  loc_0060D180: call [esi+0000001Ch]
  loc_0060D183: test eax, eax
  loc_0060D185: fnclex
  loc_0060D187: jge 0060D19Ah
  loc_0060D189: mov edx, var_AC
  loc_0060D18F: push 0000001Ch
  loc_0060D191: push 00439DD8h
  loc_0060D196: push edx
  loc_0060D197: push eax
  loc_0060D198: call edi
  loc_0060D19A: mov eax, var_28
  loc_0060D19D: lea edx, var_2C
  loc_0060D1A0: push edx
  loc_0060D1A1: push eax
  loc_0060D1A2: mov ecx, [eax]
  loc_0060D1A4: mov esi, eax
  loc_0060D1A6: call [ecx+0000003Ch]
  loc_0060D1A9: test eax, eax
  loc_0060D1AB: fnclex
  loc_0060D1AD: jge 0060D1BAh
  loc_0060D1AF: push 0000003Ch
  loc_0060D1B1: push 00439BF8h
  loc_0060D1B6: push esi
  loc_0060D1B7: push eax
  loc_0060D1B8: call edi
  loc_0060D1BA: lea ebx, var_30
  loc_0060D1BD: mov eax, var_2C
  loc_0060D1C0: push ebx
  loc_0060D1C1: mov edx, 00000008h
  loc_0060D1C6: sub esp, 00000010h
  loc_0060D1C9: mov esi, [eax]
  loc_0060D1CB: mov ebx, esp
  loc_0060D1CD: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060D1D2: push eax
  loc_0060D1D3: mov var_BC, eax
  loc_0060D1D9: mov [ebx], edx
  loc_0060D1DB: mov edx, var_7C
  loc_0060D1DE: mov [ebx+00000004h], edx
  loc_0060D1E1: mov [ebx+00000008h], ecx
  loc_0060D1E4: mov ecx, var_74
  loc_0060D1E7: mov [ebx+0000000Ch], ecx
  loc_0060D1EA: call [esi+0000001Ch]
  loc_0060D1ED: test eax, eax
  loc_0060D1EF: fnclex
  loc_0060D1F1: jge 0060D204h
  loc_0060D1F3: mov edx, var_BC
  loc_0060D1F9: push 0000001Ch
  loc_0060D1FB: push 0043B7D8h
  loc_0060D200: push edx
  loc_0060D201: push eax
  loc_0060D202: call edi
  loc_0060D204: mov ecx, [006680F8h]
  loc_0060D20A: mov edx, [006680FCh]
  loc_0060D210: sub esp, 00000010h
  loc_0060D213: mov eax, esp
  loc_0060D215: push 0043B85Ch ; "Caption"
  loc_0060D21A: mov [eax], ecx
  loc_0060D21C: mov ecx, [00668100h]
  loc_0060D222: mov [eax+00000004h], edx
  loc_0060D225: mov edx, [00668104h]
  loc_0060D22B: mov [eax+00000008h], ecx
  loc_0060D22E: mov [eax+0000000Ch], edx
  loc_0060D231: mov eax, var_30
  loc_0060D234: push eax
  loc_0060D235: call [00401094h] ; __vbaLateMemSt
  loc_0060D23B: lea ecx, var_30
  loc_0060D23E: lea edx, var_2C
  loc_0060D241: push ecx
  loc_0060D242: lea eax, var_28
  loc_0060D245: push edx
  loc_0060D246: lea ecx, var_24
  loc_0060D249: push eax
  loc_0060D24A: push ecx
  loc_0060D24B: push 00000004h
  loc_0060D24D: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D253: mov eax, var_C4
  loc_0060D259: add esp, 00000014h
  loc_0060D25C: lea ecx, var_24
  loc_0060D25F: mov edx, [eax]
  loc_0060D261: push ecx
  loc_0060D262: push eax
  loc_0060D263: call [edx+0000008Ch]
  loc_0060D269: test eax, eax
  loc_0060D26B: fnclex
  loc_0060D26D: jge 0060D283h
  loc_0060D26F: mov edx, var_C4
  loc_0060D275: push 0000008Ch
  loc_0060D27A: push 0043B5D0h
  loc_0060D27F: push edx
  loc_0060D280: push eax
  loc_0060D281: call edi
  loc_0060D283: lea ebx, var_28
  loc_0060D286: mov eax, var_24
  loc_0060D289: push ebx
  loc_0060D28A: mov edx, 00000008h
  loc_0060D28F: sub esp, 00000010h
  loc_0060D292: mov var_70, edx
  loc_0060D295: mov ebx, esp
  loc_0060D297: mov ecx, 0043B848h ; "Section4"
  loc_0060D29C: mov var_68, ecx
  loc_0060D29F: mov esi, [eax]
  loc_0060D2A1: mov [ebx], edx
  loc_0060D2A3: mov edx, var_6C
  loc_0060D2A6: push eax
  loc_0060D2A7: mov var_AC, eax
  loc_0060D2AD: mov [ebx+00000004h], edx
  loc_0060D2B0: mov [ebx+00000008h], ecx
  loc_0060D2B3: mov ecx, var_64
  loc_0060D2B6: mov [ebx+0000000Ch], ecx
  loc_0060D2B9: call [esi+0000001Ch]
  loc_0060D2BC: test eax, eax
  loc_0060D2BE: fnclex
  loc_0060D2C0: jge 0060D2D3h
  loc_0060D2C2: mov edx, var_AC
  loc_0060D2C8: push 0000001Ch
  loc_0060D2CA: push 00439DD8h
  loc_0060D2CF: push edx
  loc_0060D2D0: push eax
  loc_0060D2D1: call edi
  loc_0060D2D3: mov eax, var_28
  loc_0060D2D6: lea edx, var_2C
  loc_0060D2D9: push edx
  loc_0060D2DA: push eax
  loc_0060D2DB: mov ecx, [eax]
  loc_0060D2DD: mov esi, eax
  loc_0060D2DF: call [ecx+0000003Ch]
  loc_0060D2E2: test eax, eax
  loc_0060D2E4: fnclex
  loc_0060D2E6: jge 0060D2F3h
  loc_0060D2E8: push 0000003Ch
  loc_0060D2EA: push 00439BF8h
  loc_0060D2EF: push esi
  loc_0060D2F0: push eax
  loc_0060D2F1: call edi
  loc_0060D2F3: lea ebx, var_30
  loc_0060D2F6: mov eax, var_2C
  loc_0060D2F9: push ebx
  loc_0060D2FA: mov edx, 00000008h
  loc_0060D2FF: sub esp, 00000010h
  loc_0060D302: mov esi, [eax]
  loc_0060D304: mov ebx, esp
  loc_0060D306: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060D30B: push eax
  loc_0060D30C: mov var_BC, eax
  loc_0060D312: mov [ebx], edx
  loc_0060D314: mov edx, var_7C
  loc_0060D317: mov [ebx+00000004h], edx
  loc_0060D31A: mov [ebx+00000008h], ecx
  loc_0060D31D: mov ecx, var_74
  loc_0060D320: mov [ebx+0000000Ch], ecx
  loc_0060D323: call [esi+0000001Ch]
  loc_0060D326: test eax, eax
  loc_0060D328: fnclex
  loc_0060D32A: jge 0060D33Dh
  loc_0060D32C: mov edx, var_BC
  loc_0060D332: push 0000001Ch
  loc_0060D334: push 0043B7D8h
  loc_0060D339: push edx
  loc_0060D33A: push eax
  loc_0060D33B: call edi
  loc_0060D33D: mov ecx, [00668108h]
  loc_0060D343: mov edx, [0066810Ch]
  loc_0060D349: sub esp, 00000010h
  loc_0060D34C: mov eax, esp
  loc_0060D34E: push 0043B85Ch ; "Caption"
  loc_0060D353: mov [eax], ecx
  loc_0060D355: mov ecx, [00668110h]
  loc_0060D35B: mov [eax+00000004h], edx
  loc_0060D35E: mov edx, [00668114h]
  loc_0060D364: mov [eax+00000008h], ecx
  loc_0060D367: mov [eax+0000000Ch], edx
  loc_0060D36A: mov eax, var_30
  loc_0060D36D: push eax
  loc_0060D36E: call [00401094h] ; __vbaLateMemSt
  loc_0060D374: lea ecx, var_30
  loc_0060D377: lea edx, var_2C
  loc_0060D37A: push ecx
  loc_0060D37B: lea eax, var_28
  loc_0060D37E: push edx
  loc_0060D37F: lea ecx, var_24
  loc_0060D382: push eax
  loc_0060D383: push ecx
  loc_0060D384: push 00000004h
  loc_0060D386: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D38C: mov eax, var_C4
  loc_0060D392: add esp, 00000014h
  loc_0060D395: mov edx, [eax]
  loc_0060D397: push 0000000Ah
  loc_0060D399: push eax
  loc_0060D39A: call [edx+00000048h]
  loc_0060D39D: test eax, eax
  loc_0060D39F: fnclex
  loc_0060D3A1: jge 0060D3B4h
  loc_0060D3A3: mov ecx, var_C4
  loc_0060D3A9: push 00000048h
  loc_0060D3AB: push 0043B5D0h
  loc_0060D3B0: push ecx
  loc_0060D3B1: push eax
  loc_0060D3B2: call edi
  loc_0060D3B4: mov eax, var_C4
  loc_0060D3BA: push 0000000Ah
  loc_0060D3BC: push eax
  loc_0060D3BD: mov edx, [eax]
  loc_0060D3BF: call [edx+00000058h]
  loc_0060D3C2: test eax, eax
  loc_0060D3C4: fnclex
  loc_0060D3C6: jge 0060D3D9h
  loc_0060D3C8: mov ecx, var_C4
  loc_0060D3CE: push 00000058h
  loc_0060D3D0: push 0043B5D0h
  loc_0060D3D5: push ecx
  loc_0060D3D6: push eax
  loc_0060D3D7: call edi
  loc_0060D3D9: sub esp, 00000010h
  loc_0060D3DC: mov eax, 00000002h
  loc_0060D3E1: mov edx, esp
  loc_0060D3E3: mov ecx, eax
  loc_0060D3E5: mov var_70, ecx
  loc_0060D3E8: mov var_68, eax
  loc_0060D3EB: mov [edx], ecx
  loc_0060D3ED: mov ecx, var_6C
  loc_0060D3F0: push 8001000Eh
  loc_0060D3F5: mov [edx+00000004h], ecx
  loc_0060D3F8: mov ecx, var_C4
  loc_0060D3FE: push ecx
  loc_0060D3FF: mov [edx+00000008h], eax
  loc_0060D402: mov eax, var_64
  loc_0060D405: mov [edx+0000000Ch], eax
  loc_0060D408: call [00401280h] ; __vbaLateIdSt
  loc_0060D40E: mov edx, var_C4
  loc_0060D414: push 00000000h
  loc_0060D416: push 80011003h
  loc_0060D41B: push edx
  loc_0060D41C: call [0040102Ch] ; __vbaLateIdCall
  loc_0060D422: add esp, 0000000Ch
  loc_0060D425: lea eax, var_C4
  loc_0060D42B: push 00000000h
  loc_0060D42D: push eax
  loc_0060D42E: call [004010C8h] ; __vbaObjSetAddref
  loc_0060D434: push 0060D49Eh
  loc_0060D439: jmp 0060D474h
  loc_0060D43B: lea ecx, var_20
  loc_0060D43E: call [0040129Ch] ; __vbaFreeStr
  loc_0060D444: lea ecx, var_30
  loc_0060D447: lea edx, var_2C
  loc_0060D44A: push ecx
  loc_0060D44B: lea eax, var_28
  loc_0060D44E: push edx
  loc_0060D44F: lea ecx, var_24
  loc_0060D452: push eax
  loc_0060D453: push ecx
  loc_0060D454: push 00000004h
  loc_0060D456: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D45C: lea edx, var_60
  loc_0060D45F: lea eax, var_50
  loc_0060D462: push edx
  loc_0060D463: lea ecx, var_40
  loc_0060D466: push eax
  loc_0060D467: push ecx
  loc_0060D468: push 00000003h
  loc_0060D46A: call [0040103Ch] ; __vbaFreeVarList
  loc_0060D470: add esp, 00000024h
  loc_0060D473: ret
  loc_0060D474: lea edx, var_C8
  loc_0060D47A: lea eax, var_C4
  loc_0060D480: push edx
  loc_0060D481: push eax
  loc_0060D482: push 00000002h
  loc_0060D484: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D48A: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060D490: add esp, 0000000Ch
  loc_0060D493: lea ecx, var_18
  loc_0060D496: call __vbaFreeStr
  loc_0060D498: lea ecx, var_1C
  loc_0060D49B: call __vbaFreeStr
  loc_0060D49D: ret
  loc_0060D49E: mov ecx, var_10
  loc_0060D4A1: pop edi
  loc_0060D4A2: pop esi
  loc_0060D4A3: mov fs:[00000000h], ecx
  loc_0060D4AA: pop ebx
  loc_0060D4AB: mov esp, ebp
  loc_0060D4AD: pop ebp
  loc_0060D4AE: ret
End Sub

Private Sub Proc_34_4_60D4C0
  loc_0060D4C0: push ebp
  loc_0060D4C1: mov ebp, esp
  loc_0060D4C3: sub esp, 00000008h
  loc_0060D4C6: push 00405696h ; __vbaExceptHandler
  loc_0060D4CB: mov eax, fs:[00000000h]
  loc_0060D4D1: push eax
  loc_0060D4D2: mov fs:[00000000h], esp
  loc_0060D4D9: sub esp, 000000D0h
  loc_0060D4DF: push ebx
  loc_0060D4E0: push esi
  loc_0060D4E1: push edi
  loc_0060D4E2: mov var_8, esp
  loc_0060D4E5: mov var_4, 004044E0h
  loc_0060D4EC: xor ebx, ebx
  loc_0060D4EE: mov var_18, ebx
  loc_0060D4F1: mov var_1C, ebx
  loc_0060D4F4: mov var_20, ebx
  loc_0060D4F7: mov var_24, ebx
  loc_0060D4FA: mov var_28, ebx
  loc_0060D4FD: mov var_2C, ebx
  loc_0060D500: mov var_30, ebx
  loc_0060D503: mov var_40, ebx
  loc_0060D506: mov var_50, ebx
  loc_0060D509: mov var_60, ebx
  loc_0060D50C: mov var_70, ebx
  loc_0060D50F: mov var_80, ebx
  loc_0060D512: mov var_90, ebx
  loc_0060D518: mov var_A4, ebx
  loc_0060D51E: mov var_C4, ebx
  loc_0060D524: mov var_C8, ebx
  loc_0060D52A: call 0055B320h
  loc_0060D52F: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0060D535: mov edx, 0042DFECh
  loc_0060D53A: lea ecx, var_18
  loc_0060D53D: call __vbaStrCopy
  loc_0060D53F: mov edx, 0043BB68h ; " SELECT kode_jenis, nama_jenis FROM brg_jenis ORDER BY kode_jenis"
  loc_0060D544: lea ecx, var_18
  loc_0060D547: call __vbaStrCopy
  loc_0060D549: push 0042DF30h
  loc_0060D54E: call [00401168h] ; __vbaNew
  loc_0060D554: push eax
  loc_0060D555: push 00668090h
  loc_0060D55A: call [004010B8h] ; __vbaObjSet
  loc_0060D560: cmp [0066803Ch], ebx
  loc_0060D566: jnz 0060D578h
  loc_0060D568: push 0066803Ch
  loc_0060D56D: push 0042DEFCh
  loc_0060D572: call [004011E8h] ; __vbaNew2
  loc_0060D578: mov esi, [0066803Ch]
  loc_0060D57E: lea ecx, var_40
  loc_0060D581: mov var_38, 80020004h
  loc_0060D588: mov var_40, 0000000Ah
  loc_0060D58F: call [0040123Ch] ; __vbaFreeVarg
  loc_0060D595: mov eax, [esi]
  loc_0060D597: lea ecx, var_24
  loc_0060D59A: push ecx
  loc_0060D59B: mov ecx, var_18
  loc_0060D59E: lea edx, var_40
  loc_0060D5A1: push 00000001h
  loc_0060D5A3: push edx
  loc_0060D5A4: push ecx
  loc_0060D5A5: push esi
  loc_0060D5A6: call [eax+00000040h]
  loc_0060D5A9: cmp eax, ebx
  loc_0060D5AB: fnclex
  loc_0060D5AD: jge 0060D5C2h
  loc_0060D5AF: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060D5B5: push 00000040h
  loc_0060D5B7: push 0042E1B0h
  loc_0060D5BC: push esi
  loc_0060D5BD: push eax
  loc_0060D5BE: call edi
  loc_0060D5C0: jmp 0060D5C8h
  loc_0060D5C2: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060D5C8: mov edx, var_24
  loc_0060D5CB: mov esi, [00401268h] ; __vbaCastObj
  loc_0060D5D1: push 0042DF20h
  loc_0060D5D6: push edx
  loc_0060D5D7: call __vbaCastObj
  loc_0060D5D9: push eax
  loc_0060D5DA: push 00668090h
  loc_0060D5DF: call [004010B8h] ; __vbaObjSet
  loc_0060D5E5: lea ecx, var_24
  loc_0060D5E8: call [004012A0h] ; __vbaFreeObj
  loc_0060D5EE: lea ecx, var_40
  loc_0060D5F1: call [00401020h] ; __vbaFreeVar
  loc_0060D5F7: cmp [00668544h], ebx
  loc_0060D5FD: jnz 0060D60Fh
  loc_0060D5FF: push 00668544h
  loc_0060D604: push 004081CCh
  loc_0060D609: call [004011E8h] ; __vbaNew2
  loc_0060D60F: mov eax, [00668544h]
  loc_0060D614: lea ecx, var_C4
  loc_0060D61A: push eax
  loc_0060D61B: push ecx
  loc_0060D61C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060D622: mov edx, var_C4
  loc_0060D628: push 0043B7B8h
  loc_0060D62D: push ebx
  loc_0060D62E: mov edx, [edx]
  loc_0060D630: mov var_E0, edx
  loc_0060D636: call __vbaCastObj
  loc_0060D638: push eax
  loc_0060D639: lea eax, var_24
  loc_0060D63C: push eax
  loc_0060D63D: call [004010B8h] ; __vbaObjSet
  loc_0060D643: mov ecx, var_C4
  loc_0060D649: mov edx, var_E0
  loc_0060D64F: push eax
  loc_0060D650: push ecx
  loc_0060D651: call [edx+00000028h]
  loc_0060D654: cmp eax, ebx
  loc_0060D656: fnclex
  loc_0060D658: jge 0060D66Bh
  loc_0060D65A: mov ecx, var_C4
  loc_0060D660: push 00000028h
  loc_0060D662: push 0043B5D0h
  loc_0060D667: push ecx
  loc_0060D668: push eax
  loc_0060D669: call edi
  loc_0060D66B: lea ecx, var_24
  loc_0060D66E: call [004012A0h] ; __vbaFreeObj
  loc_0060D674: mov eax, var_C4
  loc_0060D67A: push 0042DFECh
  loc_0060D67F: push eax
  loc_0060D680: mov edx, [eax]
  loc_0060D682: call [edx+00000030h]
  loc_0060D685: cmp eax, ebx
  loc_0060D687: fnclex
  loc_0060D689: jge 0060D69Ch
  loc_0060D68B: mov ecx, var_C4
  loc_0060D691: push 00000030h
  loc_0060D693: push 0043B5D0h
  loc_0060D698: push ecx
  loc_0060D699: push eax
  loc_0060D69A: call edi
  loc_0060D69C: mov eax, [00668090h]
  loc_0060D6A1: lea ecx, var_24
  loc_0060D6A4: push ecx
  loc_0060D6A5: push eax
  loc_0060D6A6: mov edx, [eax]
  loc_0060D6A8: call [edx+00000114h]
  loc_0060D6AE: cmp eax, ebx
  loc_0060D6B0: fnclex
  loc_0060D6B2: jge 0060D6C8h
  loc_0060D6B4: mov edx, [00668090h]
  loc_0060D6BA: push 00000114h
  loc_0060D6BF: push 0043B7C8h
  loc_0060D6C4: push edx
  loc_0060D6C5: push eax
  loc_0060D6C6: call edi
  loc_0060D6C8: mov eax, var_C4
  loc_0060D6CE: mov ecx, var_24
  loc_0060D6D1: push 0043B7B8h
  loc_0060D6D6: push ecx
  loc_0060D6D7: mov esi, [eax]
  loc_0060D6D9: call [00401268h] ; __vbaCastObj
  loc_0060D6DF: lea edx, var_28
  loc_0060D6E2: push eax
  loc_0060D6E3: push edx
  loc_0060D6E4: call [004010B8h] ; __vbaObjSet
  loc_0060D6EA: push eax
  loc_0060D6EB: mov eax, var_C4
  loc_0060D6F1: push eax
  loc_0060D6F2: call [esi+00000028h]
  loc_0060D6F5: cmp eax, ebx
  loc_0060D6F7: fnclex
  loc_0060D6F9: jge 0060D70Ch
  loc_0060D6FB: mov ecx, var_C4
  loc_0060D701: push 00000028h
  loc_0060D703: push 0043B5D0h
  loc_0060D708: push ecx
  loc_0060D709: push eax
  loc_0060D70A: call edi
  loc_0060D70C: lea edx, var_28
  loc_0060D70F: lea eax, var_24
  loc_0060D712: push edx
  loc_0060D713: push eax
  loc_0060D714: push 00000002h
  loc_0060D716: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D71C: mov eax, var_C4
  loc_0060D722: add esp, 0000000Ch
  loc_0060D725: lea edx, var_24
  loc_0060D728: mov ecx, [eax]
  loc_0060D72A: push edx
  loc_0060D72B: push eax
  loc_0060D72C: call [ecx+0000008Ch]
  loc_0060D732: cmp eax, ebx
  loc_0060D734: fnclex
  loc_0060D736: jge 0060D74Ch
  loc_0060D738: mov ecx, var_C4
  loc_0060D73E: push 0000008Ch
  loc_0060D743: push 0043B5D0h
  loc_0060D748: push ecx
  loc_0060D749: push eax
  loc_0060D74A: call edi
  loc_0060D74C: lea esi, var_28
  loc_0060D74F: mov eax, var_24
  loc_0060D752: push esi
  loc_0060D753: mov ecx, 00000008h
  loc_0060D758: sub esp, 00000010h
  loc_0060D75B: mov var_70, ecx
  loc_0060D75E: mov esi, esp
  loc_0060D760: mov var_68, 00439FB0h ; "Section1"
  loc_0060D767: mov edx, [eax]
  loc_0060D769: push eax
  loc_0060D76A: mov [esi], ecx
  loc_0060D76C: mov ecx, var_6C
  loc_0060D76F: mov var_AC, eax
  loc_0060D775: mov [esi+00000004h], ecx
  loc_0060D778: mov ecx, var_68
  loc_0060D77B: mov [esi+00000008h], ecx
  loc_0060D77E: mov ecx, var_64
  loc_0060D781: mov [esi+0000000Ch], ecx
  loc_0060D784: call [edx+0000001Ch]
  loc_0060D787: cmp eax, ebx
  loc_0060D789: fnclex
  loc_0060D78B: jge 0060D79Eh
  loc_0060D78D: mov edx, var_AC
  loc_0060D793: push 0000001Ch
  loc_0060D795: push 00439DD8h
  loc_0060D79A: push edx
  loc_0060D79B: push eax
  loc_0060D79C: call edi
  loc_0060D79E: mov eax, var_28
  loc_0060D7A1: lea edx, var_2C
  loc_0060D7A4: push edx
  loc_0060D7A5: push eax
  loc_0060D7A6: mov ecx, [eax]
  loc_0060D7A8: mov esi, eax
  loc_0060D7AA: call [ecx+0000003Ch]
  loc_0060D7AD: cmp eax, ebx
  loc_0060D7AF: fnclex
  loc_0060D7B1: jge 0060D7BEh
  loc_0060D7B3: push 0000003Ch
  loc_0060D7B5: push 00439BF8h
  loc_0060D7BA: push esi
  loc_0060D7BB: push eax
  loc_0060D7BC: call edi
  loc_0060D7BE: mov eax, var_2C
  loc_0060D7C1: mov var_2C, ebx
  loc_0060D7C4: push eax
  loc_0060D7C5: lea eax, var_C8
  loc_0060D7CB: push eax
  loc_0060D7CC: call [004010B8h] ; __vbaObjSet
  loc_0060D7D2: lea ecx, var_28
  loc_0060D7D5: lea edx, var_24
  loc_0060D7D8: push ecx
  loc_0060D7D9: push edx
  loc_0060D7DA: push 00000002h
  loc_0060D7DC: call [0040104Ch] ; __vbaFreeObjList
  loc_0060D7E2: mov eax, var_C8
  loc_0060D7E8: add esp, 0000000Ch
  loc_0060D7EB: lea edx, var_A4
  loc_0060D7F1: mov ecx, [eax]
  loc_0060D7F3: push edx
  loc_0060D7F4: push eax
  loc_0060D7F5: call [ecx+00000020h]
  loc_0060D7F8: cmp eax, ebx
  loc_0060D7FA: fnclex
  loc_0060D7FC: jge 0060D80Fh
  loc_0060D7FE: mov ecx, var_C8
  loc_0060D804: push 00000020h
  loc_0060D806: push 0043B7D8h
  loc_0060D80B: push ecx
  loc_0060D80C: push eax
  loc_0060D80D: call edi
  loc_0060D80F: mov ecx, var_A4
  loc_0060D815: call [00401138h] ; __vbaI2I4
  loc_0060D81B: mov var_D0, eax
  loc_0060D821: mov ebx, 00000001h
  loc_0060D826: cmp bx, var_D0
  loc_0060D82D: mov var_14, ebx
  loc_0060D830: jg 0060DAA1h
  loc_0060D836: lea esi, var_24
  loc_0060D839: mov ecx, var_C8
  loc_0060D83F: push esi
  loc_0060D840: mov eax, 00000002h
  loc_0060D845: sub esp, 00000010h
  loc_0060D848: mov var_70, eax
  loc_0060D84B: mov esi, esp
  loc_0060D84D: mov var_68, bx
  loc_0060D851: mov edx, [ecx]
  loc_0060D853: push ecx
  loc_0060D854: mov [esi], eax
  loc_0060D856: mov eax, var_6C
  loc_0060D859: mov [esi+00000004h], eax
  loc_0060D85C: mov eax, var_68
  loc_0060D85F: mov [esi+00000008h], eax
  loc_0060D862: mov eax, var_64
  loc_0060D865: mov [esi+0000000Ch], eax
  loc_0060D868: call [edx+0000001Ch]
  loc_0060D86B: test eax, eax
  loc_0060D86D: fnclex
  loc_0060D86F: jge 0060D882h
  loc_0060D871: mov ecx, var_C8
  loc_0060D877: push 0000001Ch
  loc_0060D879: push 0043B7D8h
  loc_0060D87E: push ecx
  loc_0060D87F: push eax
  loc_0060D880: call edi
  loc_0060D882: mov edx, var_24
  loc_0060D885: push 0043B7E8h
  loc_0060D88A: push edx
  loc_0060D88B: call [004011CCh] ; __vbaCheckType
  loc_0060D891: lea ecx, var_24
  loc_0060D894: mov si, ax
  loc_0060D897: call [004012A0h] ; __vbaFreeObj
  loc_0060D89D: test si, si
  loc_0060D8A0: Unknown_760()
  loc_0060D8A6: mov var_68, bx
  loc_0060D8AA: lea ebx, var_24
  loc_0060D8AD: push ebx
  loc_0060D8AE: mov ecx, var_C8
  loc_0060D8B4: sub esp, 00000010h
  loc_0060D8B7: mov eax, 00000002h
  loc_0060D8BC: mov ebx, esp
  loc_0060D8BE: mov var_70, eax
  loc_0060D8C1: mov edx, [ecx]
  loc_0060D8C3: mov esi, 0042DFECh
  loc_0060D8C8: mov [ebx], eax
  loc_0060D8CA: mov eax, var_6C
  loc_0060D8CD: push ecx
  loc_0060D8CE: mov var_78, esi
  loc_0060D8D1: mov [ebx+00000004h], eax
  loc_0060D8D4: mov eax, var_68
  loc_0060D8D7: mov edi, 00000008h
  loc_0060D8DC: mov [ebx+00000008h], eax
  loc_0060D8DF: mov eax, var_64
  loc_0060D8E2: mov [ebx+0000000Ch], eax
  loc_0060D8E5: call [edx+0000001Ch]
  loc_0060D8E8: test eax, eax
  loc_0060D8EA: fnclex
  loc_0060D8EC: jge 0060D903h
  loc_0060D8EE: mov ecx, var_C8
  loc_0060D8F4: push 0000001Ch
  loc_0060D8F6: push 0043B7D8h
  loc_0060D8FB: push ecx
  loc_0060D8FC: push eax
  loc_0060D8FD: call [00401080h] ; __vbaHresultCheckObj
  loc_0060D903: mov eax, var_7C
  loc_0060D906: sub esp, 00000010h
  loc_0060D909: mov edx, esp
  loc_0060D90B: mov ecx, var_74
  loc_0060D90E: push 0043B7F8h ; "DataMember"
  loc_0060D913: mov [edx], edi
  loc_0060D915: mov [edx+00000004h], eax
  loc_0060D918: mov [edx+00000008h], esi
  loc_0060D91B: mov [edx+0000000Ch], ecx
  loc_0060D91E: mov edx, var_24
  loc_0060D921: push edx
  loc_0060D922: call [00401094h] ; __vbaLateMemSt
  loc_0060D928: lea ecx, var_24
  loc_0060D92B: call [004012A0h] ; __vbaFreeObj
  loc_0060D931: mov eax, [00668090h]
  loc_0060D936: lea edx, var_24
  loc_0060D939: push edx
  loc_0060D93A: push eax
  loc_0060D93B: mov ecx, [eax]
  loc_0060D93D: call [ecx+00000054h]
  loc_0060D940: test eax, eax
  loc_0060D942: fnclex
  loc_0060D944: jge 0060D95Bh
  loc_0060D946: mov ecx, [00668090h]
  loc_0060D94C: push 00000054h
  loc_0060D94E: push 0042DF44h
  loc_0060D953: push ecx
  loc_0060D954: push eax
  loc_0060D955: call [00401080h] ; __vbaHresultCheckObj
  loc_0060D95B: mov ebx, var_14
  loc_0060D95E: lea edi, var_28
  loc_0060D961: mov dx, bx
  loc_0060D964: push edi
  loc_0060D965: sub dx, 0001h
  loc_0060D969: mov eax, var_24
  loc_0060D96C: jo 0060E1FAh
  loc_0060D972: sub esp, 00000010h
  loc_0060D975: mov ecx, 00000002h
  loc_0060D97A: mov edi, esp
  loc_0060D97C: mov var_70, ecx
  loc_0060D97F: mov var_68, dx
  loc_0060D983: mov edx, [eax]
  loc_0060D985: mov [edi], ecx
  loc_0060D987: mov ecx, var_6C
  loc_0060D98A: push eax
  loc_0060D98B: mov esi, eax
  loc_0060D98D: mov [edi+00000004h], ecx
  loc_0060D990: mov ecx, var_68
  loc_0060D993: mov [edi+00000008h], ecx
  loc_0060D996: mov ecx, var_64
  loc_0060D999: mov [edi+0000000Ch], ecx
  loc_0060D99C: call [edx+00000028h]
  loc_0060D99F: test eax, eax
  loc_0060D9A1: fnclex
  loc_0060D9A3: jge 0060D9B8h
  loc_0060D9A5: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060D9AB: push 00000028h
  loc_0060D9AD: push 0042DFACh
  loc_0060D9B2: push esi
  loc_0060D9B3: push eax
  loc_0060D9B4: call edi
  loc_0060D9B6: jmp 0060D9BEh
  loc_0060D9B8: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060D9BE: mov eax, var_28
  loc_0060D9C1: lea ecx, var_20
  loc_0060D9C4: push ecx
  loc_0060D9C5: push eax
  loc_0060D9C6: mov edx, [eax]
  loc_0060D9C8: mov esi, eax
  loc_0060D9CA: call [edx+0000002Ch]
  loc_0060D9CD: test eax, eax
  loc_0060D9CF: fnclex
  loc_0060D9D1: jge 0060D9DEh
  loc_0060D9D3: push 0000002Ch
  loc_0060D9D5: push 0042DFBCh
  loc_0060D9DA: push esi
  loc_0060D9DB: push eax
  loc_0060D9DC: call edi
  loc_0060D9DE: mov eax, var_20
  loc_0060D9E1: lea esi, var_2C
  loc_0060D9E4: push esi
  loc_0060D9E5: mov ecx, var_C8
  loc_0060D9EB: sub esp, 00000010h
  loc_0060D9EE: mov var_38, eax
  loc_0060D9F1: mov esi, esp
  loc_0060D9F3: mov eax, 00000002h
  loc_0060D9F8: mov var_78, bx
  loc_0060D9FC: mov var_20, 00000000h
  loc_0060DA03: mov [esi], eax
  loc_0060DA05: mov eax, var_7C
  loc_0060DA08: mov var_40, 00000008h
  loc_0060DA0F: mov edx, [ecx]
  loc_0060DA11: mov [esi+00000004h], eax
  loc_0060DA14: mov eax, var_78
  loc_0060DA17: push ecx
  loc_0060DA18: mov [esi+00000008h], eax
  loc_0060DA1B: mov eax, var_74
  loc_0060DA1E: mov [esi+0000000Ch], eax
  loc_0060DA21: call [edx+0000001Ch]
  loc_0060DA24: test eax, eax
  loc_0060DA26: fnclex
  loc_0060DA28: jge 0060DA3Bh
  loc_0060DA2A: mov ecx, var_C8
  loc_0060DA30: push 0000001Ch
  loc_0060DA32: push 0043B7D8h
  loc_0060DA37: push ecx
  loc_0060DA38: push eax
  loc_0060DA39: call edi
  loc_0060DA3B: mov eax, var_40
  loc_0060DA3E: mov ecx, var_3C
  loc_0060DA41: sub esp, 00000010h
  loc_0060DA44: mov edx, esp
  loc_0060DA46: push 0043B810h ; "DataField"
  loc_0060DA4B: mov [edx], eax
  loc_0060DA4D: mov eax, var_38
  loc_0060DA50: mov [edx+00000004h], ecx
  loc_0060DA53: mov ecx, var_34
  loc_0060DA56: mov [edx+00000008h], eax
  loc_0060DA59: mov [edx+0000000Ch], ecx
  loc_0060DA5C: mov edx, var_2C
  loc_0060DA5F: push edx
  loc_0060DA60: call [00401094h] ; __vbaLateMemSt
  loc_0060DA66: lea eax, var_2C
  loc_0060DA69: lea ecx, var_28
  loc_0060DA6C: push eax
  loc_0060DA6D: lea edx, var_24
  loc_0060DA70: push ecx
  loc_0060DA71: push edx
  loc_0060DA72: push 00000003h
  loc_0060DA74: call [0040104Ch] ; __vbaFreeObjList
  loc_0060DA7A: add esp, 00000010h
  loc_0060DA7D: lea ecx, var_40
  loc_0060DA80: call [00401020h] ; __vbaFreeVar
  loc_0060DA86: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060DA8C: mov eax, 00000001h
  loc_0060DA91: add ax, bx
  loc_0060DA94: jo 0060E1FAh
  loc_0060DA9A: mov ebx, eax
  loc_0060DA9C: jmp 0060D826h
  loc_0060DAA1: lea eax, var_C8
  loc_0060DAA7: push 00000000h
  loc_0060DAA9: push eax
  loc_0060DAAA: call [004010C8h] ; __vbaObjSetAddref
  loc_0060DAB0: lea ecx, var_40
  loc_0060DAB3: push ecx
  loc_0060DAB4: call [00401220h] ; rtcGetDateVar
  loc_0060DABA: lea edx, var_70
  loc_0060DABD: lea ecx, var_50
  loc_0060DAC0: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060DAC7: mov var_70, 00000008h
  loc_0060DACE: call [00401238h] ; __vbaVarDup
  loc_0060DAD4: push 00000001h
  loc_0060DAD6: lea edx, var_50
  loc_0060DAD9: push 00000001h
  loc_0060DADB: lea eax, var_40
  loc_0060DADE: push edx
  loc_0060DADF: lea ecx, var_60
  loc_0060DAE2: push eax
  loc_0060DAE3: push ecx
  loc_0060DAE4: call [00401078h] ; rtcVarFromFormatVar
  loc_0060DAEA: mov eax, var_C4
  loc_0060DAF0: lea ecx, var_24
  loc_0060DAF3: push ecx
  loc_0060DAF4: push eax
  loc_0060DAF5: mov edx, [eax]
  loc_0060DAF7: call [edx+0000008Ch]
  loc_0060DAFD: test eax, eax
  loc_0060DAFF: fnclex
  loc_0060DB01: jge 0060DB17h
  loc_0060DB03: mov edx, var_C4
  loc_0060DB09: push 0000008Ch
  loc_0060DB0E: push 0043B5D0h
  loc_0060DB13: push edx
  loc_0060DB14: push eax
  loc_0060DB15: call edi
  loc_0060DB17: lea ebx, var_28
  loc_0060DB1A: mov eax, var_24
  loc_0060DB1D: push ebx
  loc_0060DB1E: mov edx, 00000008h
  loc_0060DB23: sub esp, 00000010h
  loc_0060DB26: mov esi, [eax]
  loc_0060DB28: mov ebx, esp
  loc_0060DB2A: mov ecx, 0043B848h ; "Section4"
  loc_0060DB2F: push eax
  loc_0060DB30: mov var_AC, eax
  loc_0060DB36: mov [ebx], edx
  loc_0060DB38: mov edx, var_7C
  loc_0060DB3B: mov [ebx+00000004h], edx
  loc_0060DB3E: mov [ebx+00000008h], ecx
  loc_0060DB41: mov ecx, var_74
  loc_0060DB44: mov [ebx+0000000Ch], ecx
  loc_0060DB47: call [esi+0000001Ch]
  loc_0060DB4A: test eax, eax
  loc_0060DB4C: fnclex
  loc_0060DB4E: jge 0060DB61h
  loc_0060DB50: mov edx, var_AC
  loc_0060DB56: push 0000001Ch
  loc_0060DB58: push 00439DD8h
  loc_0060DB5D: push edx
  loc_0060DB5E: push eax
  loc_0060DB5F: call edi
  loc_0060DB61: mov eax, var_28
  loc_0060DB64: lea edx, var_2C
  loc_0060DB67: push edx
  loc_0060DB68: push eax
  loc_0060DB69: mov ecx, [eax]
  loc_0060DB6B: mov esi, eax
  loc_0060DB6D: call [ecx+0000003Ch]
  loc_0060DB70: test eax, eax
  loc_0060DB72: fnclex
  loc_0060DB74: jge 0060DB81h
  loc_0060DB76: push 0000003Ch
  loc_0060DB78: push 00439BF8h
  loc_0060DB7D: push esi
  loc_0060DB7E: push eax
  loc_0060DB7F: call edi
  loc_0060DB81: lea ebx, var_30
  loc_0060DB84: mov eax, var_2C
  loc_0060DB87: push ebx
  loc_0060DB88: mov edx, 00000008h
  loc_0060DB8D: sub esp, 00000010h
  loc_0060DB90: mov esi, [eax]
  loc_0060DB92: mov ebx, esp
  loc_0060DB94: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060DB99: push eax
  loc_0060DB9A: mov var_BC, eax
  loc_0060DBA0: mov [ebx], edx
  loc_0060DBA2: mov edx, var_8C
  loc_0060DBA8: mov [ebx+00000004h], edx
  loc_0060DBAB: mov [ebx+00000008h], ecx
  loc_0060DBAE: mov ecx, var_84
  loc_0060DBB4: mov [ebx+0000000Ch], ecx
  loc_0060DBB7: call [esi+0000001Ch]
  loc_0060DBBA: test eax, eax
  loc_0060DBBC: fnclex
  loc_0060DBBE: jge 0060DBD1h
  loc_0060DBC0: mov edx, var_BC
  loc_0060DBC6: push 0000001Ch
  loc_0060DBC8: push 0043B7D8h
  loc_0060DBCD: push edx
  loc_0060DBCE: push eax
  loc_0060DBCF: call edi
  loc_0060DBD1: mov ecx, var_60
  loc_0060DBD4: mov edx, var_5C
  loc_0060DBD7: sub esp, 00000010h
  loc_0060DBDA: mov eax, esp
  loc_0060DBDC: push 0043B85Ch ; "Caption"
  loc_0060DBE1: mov [eax], ecx
  loc_0060DBE3: mov ecx, var_58
  loc_0060DBE6: mov [eax+00000004h], edx
  loc_0060DBE9: mov edx, var_54
  loc_0060DBEC: mov [eax+00000008h], ecx
  loc_0060DBEF: mov [eax+0000000Ch], edx
  loc_0060DBF2: mov eax, var_30
  loc_0060DBF5: push eax
  loc_0060DBF6: call [00401094h] ; __vbaLateMemSt
  loc_0060DBFC: lea ecx, var_30
  loc_0060DBFF: lea edx, var_2C
  loc_0060DC02: push ecx
  loc_0060DC03: lea eax, var_28
  loc_0060DC06: push edx
  loc_0060DC07: lea ecx, var_24
  loc_0060DC0A: push eax
  loc_0060DC0B: push ecx
  loc_0060DC0C: push 00000004h
  loc_0060DC0E: call [0040104Ch] ; __vbaFreeObjList
  loc_0060DC14: lea edx, var_60
  loc_0060DC17: lea eax, var_50
  loc_0060DC1A: push edx
  loc_0060DC1B: lea ecx, var_40
  loc_0060DC1E: push eax
  loc_0060DC1F: push ecx
  loc_0060DC20: push 00000003h
  loc_0060DC22: call [0040103Ch] ; __vbaFreeVarList
  loc_0060DC28: mov eax, var_C4
  loc_0060DC2E: add esp, 00000024h
  loc_0060DC31: lea ecx, var_24
  loc_0060DC34: mov edx, [eax]
  loc_0060DC36: push ecx
  loc_0060DC37: push eax
  loc_0060DC38: call [edx+0000008Ch]
  loc_0060DC3E: test eax, eax
  loc_0060DC40: fnclex
  loc_0060DC42: jge 0060DC58h
  loc_0060DC44: mov edx, var_C4
  loc_0060DC4A: push 0000008Ch
  loc_0060DC4F: push 0043B5D0h
  loc_0060DC54: push edx
  loc_0060DC55: push eax
  loc_0060DC56: call edi
  loc_0060DC58: lea ebx, var_28
  loc_0060DC5B: mov eax, var_24
  loc_0060DC5E: push ebx
  loc_0060DC5F: mov edx, 00000008h
  loc_0060DC64: sub esp, 00000010h
  loc_0060DC67: mov var_70, edx
  loc_0060DC6A: mov ebx, esp
  loc_0060DC6C: mov ecx, 0043B848h ; "Section4"
  loc_0060DC71: mov var_68, ecx
  loc_0060DC74: mov esi, [eax]
  loc_0060DC76: mov [ebx], edx
  loc_0060DC78: mov edx, var_6C
  loc_0060DC7B: push eax
  loc_0060DC7C: mov var_AC, eax
  loc_0060DC82: mov [ebx+00000004h], edx
  loc_0060DC85: mov [ebx+00000008h], ecx
  loc_0060DC88: mov ecx, var_64
  loc_0060DC8B: mov [ebx+0000000Ch], ecx
  loc_0060DC8E: call [esi+0000001Ch]
  loc_0060DC91: test eax, eax
  loc_0060DC93: fnclex
  loc_0060DC95: jge 0060DCA8h
  loc_0060DC97: mov edx, var_AC
  loc_0060DC9D: push 0000001Ch
  loc_0060DC9F: push 00439DD8h
  loc_0060DCA4: push edx
  loc_0060DCA5: push eax
  loc_0060DCA6: call edi
  loc_0060DCA8: mov eax, var_28
  loc_0060DCAB: lea edx, var_2C
  loc_0060DCAE: push edx
  loc_0060DCAF: push eax
  loc_0060DCB0: mov ecx, [eax]
  loc_0060DCB2: mov esi, eax
  loc_0060DCB4: call [ecx+0000003Ch]
  loc_0060DCB7: test eax, eax
  loc_0060DCB9: fnclex
  loc_0060DCBB: jge 0060DCC8h
  loc_0060DCBD: push 0000003Ch
  loc_0060DCBF: push 00439BF8h
  loc_0060DCC4: push esi
  loc_0060DCC5: push eax
  loc_0060DCC6: call edi
  loc_0060DCC8: lea ebx, var_30
  loc_0060DCCB: mov eax, var_2C
  loc_0060DCCE: push ebx
  loc_0060DCCF: mov edx, 00000008h
  loc_0060DCD4: sub esp, 00000010h
  loc_0060DCD7: mov esi, [eax]
  loc_0060DCD9: mov ebx, esp
  loc_0060DCDB: mov ecx, 0043B870h ; "LabelNama"
  loc_0060DCE0: push eax
  loc_0060DCE1: mov var_BC, eax
  loc_0060DCE7: mov [ebx], edx
  loc_0060DCE9: mov edx, var_7C
  loc_0060DCEC: mov [ebx+00000004h], edx
  loc_0060DCEF: mov [ebx+00000008h], ecx
  loc_0060DCF2: mov ecx, var_74
  loc_0060DCF5: mov [ebx+0000000Ch], ecx
  loc_0060DCF8: call [esi+0000001Ch]
  loc_0060DCFB: test eax, eax
  loc_0060DCFD: fnclex
  loc_0060DCFF: jge 0060DD12h
  loc_0060DD01: mov edx, var_BC
  loc_0060DD07: push 0000001Ch
  loc_0060DD09: push 0043B7D8h
  loc_0060DD0E: push edx
  loc_0060DD0F: push eax
  loc_0060DD10: call edi
  loc_0060DD12: mov ecx, [006680D8h]
  loc_0060DD18: mov edx, [006680DCh]
  loc_0060DD1E: sub esp, 00000010h
  loc_0060DD21: mov eax, esp
  loc_0060DD23: push 0043B85Ch ; "Caption"
  loc_0060DD28: mov [eax], ecx
  loc_0060DD2A: mov ecx, [006680E0h]
  loc_0060DD30: mov [eax+00000004h], edx
  loc_0060DD33: mov edx, [006680E4h]
  loc_0060DD39: mov [eax+00000008h], ecx
  loc_0060DD3C: mov [eax+0000000Ch], edx
  loc_0060DD3F: mov eax, var_30
  loc_0060DD42: push eax
  loc_0060DD43: call [00401094h] ; __vbaLateMemSt
  loc_0060DD49: lea ecx, var_30
  loc_0060DD4C: lea edx, var_2C
  loc_0060DD4F: push ecx
  loc_0060DD50: lea eax, var_28
  loc_0060DD53: push edx
  loc_0060DD54: lea ecx, var_24
  loc_0060DD57: push eax
  loc_0060DD58: push ecx
  loc_0060DD59: push 00000004h
  loc_0060DD5B: call [0040104Ch] ; __vbaFreeObjList
  loc_0060DD61: mov eax, var_C4
  loc_0060DD67: add esp, 00000014h
  loc_0060DD6A: lea ecx, var_24
  loc_0060DD6D: mov edx, [eax]
  loc_0060DD6F: push ecx
  loc_0060DD70: push eax
  loc_0060DD71: call [edx+0000008Ch]
  loc_0060DD77: test eax, eax
  loc_0060DD79: fnclex
  loc_0060DD7B: jge 0060DD91h
  loc_0060DD7D: mov edx, var_C4
  loc_0060DD83: push 0000008Ch
  loc_0060DD88: push 0043B5D0h
  loc_0060DD8D: push edx
  loc_0060DD8E: push eax
  loc_0060DD8F: call edi
  loc_0060DD91: lea ebx, var_28
  loc_0060DD94: mov eax, var_24
  loc_0060DD97: push ebx
  loc_0060DD98: mov edx, 00000008h
  loc_0060DD9D: sub esp, 00000010h
  loc_0060DDA0: mov var_70, edx
  loc_0060DDA3: mov ebx, esp
  loc_0060DDA5: mov ecx, 0043B848h ; "Section4"
  loc_0060DDAA: mov var_68, ecx
  loc_0060DDAD: mov esi, [eax]
  loc_0060DDAF: mov [ebx], edx
  loc_0060DDB1: mov edx, var_6C
  loc_0060DDB4: push eax
  loc_0060DDB5: mov var_AC, eax
  loc_0060DDBB: mov [ebx+00000004h], edx
  loc_0060DDBE: mov [ebx+00000008h], ecx
  loc_0060DDC1: mov ecx, var_64
  loc_0060DDC4: mov [ebx+0000000Ch], ecx
  loc_0060DDC7: call [esi+0000001Ch]
  loc_0060DDCA: test eax, eax
  loc_0060DDCC: fnclex
  loc_0060DDCE: jge 0060DDE1h
  loc_0060DDD0: mov edx, var_AC
  loc_0060DDD6: push 0000001Ch
  loc_0060DDD8: push 00439DD8h
  loc_0060DDDD: push edx
  loc_0060DDDE: push eax
  loc_0060DDDF: call edi
  loc_0060DDE1: mov eax, var_28
  loc_0060DDE4: lea edx, var_2C
  loc_0060DDE7: push edx
  loc_0060DDE8: push eax
  loc_0060DDE9: mov ecx, [eax]
  loc_0060DDEB: mov esi, eax
  loc_0060DDED: call [ecx+0000003Ch]
  loc_0060DDF0: test eax, eax
  loc_0060DDF2: fnclex
  loc_0060DDF4: jge 0060DE01h
  loc_0060DDF6: push 0000003Ch
  loc_0060DDF8: push 00439BF8h
  loc_0060DDFD: push esi
  loc_0060DDFE: push eax
  loc_0060DDFF: call edi
  loc_0060DE01: lea ebx, var_30
  loc_0060DE04: mov eax, var_2C
  loc_0060DE07: push ebx
  loc_0060DE08: mov edx, 00000008h
  loc_0060DE0D: sub esp, 00000010h
  loc_0060DE10: mov esi, [eax]
  loc_0060DE12: mov ebx, esp
  loc_0060DE14: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060DE19: push eax
  loc_0060DE1A: mov var_BC, eax
  loc_0060DE20: mov [ebx], edx
  loc_0060DE22: mov edx, var_7C
  loc_0060DE25: mov [ebx+00000004h], edx
  loc_0060DE28: mov [ebx+00000008h], ecx
  loc_0060DE2B: mov ecx, var_74
  loc_0060DE2E: mov [ebx+0000000Ch], ecx
  loc_0060DE31: call [esi+0000001Ch]
  loc_0060DE34: test eax, eax
  loc_0060DE36: fnclex
  loc_0060DE38: jge 0060DE4Bh
  loc_0060DE3A: mov edx, var_BC
  loc_0060DE40: push 0000001Ch
  loc_0060DE42: push 0043B7D8h
  loc_0060DE47: push edx
  loc_0060DE48: push eax
  loc_0060DE49: call edi
  loc_0060DE4B: mov ecx, [006680E8h]
  loc_0060DE51: mov edx, [006680ECh]
  loc_0060DE57: sub esp, 00000010h
  loc_0060DE5A: mov eax, esp
  loc_0060DE5C: push 0043B85Ch ; "Caption"
  loc_0060DE61: mov [eax], ecx
  loc_0060DE63: mov ecx, [006680F0h]
  loc_0060DE69: mov [eax+00000004h], edx
  loc_0060DE6C: mov edx, [006680F4h]
  loc_0060DE72: mov [eax+00000008h], ecx
  loc_0060DE75: mov [eax+0000000Ch], edx
  loc_0060DE78: mov eax, var_30
  loc_0060DE7B: push eax
  loc_0060DE7C: call [00401094h] ; __vbaLateMemSt
  loc_0060DE82: lea ecx, var_30
  loc_0060DE85: lea edx, var_2C
  loc_0060DE88: push ecx
  loc_0060DE89: lea eax, var_28
  loc_0060DE8C: push edx
  loc_0060DE8D: lea ecx, var_24
  loc_0060DE90: push eax
  loc_0060DE91: push ecx
  loc_0060DE92: push 00000004h
  loc_0060DE94: call [0040104Ch] ; __vbaFreeObjList
  loc_0060DE9A: mov eax, var_C4
  loc_0060DEA0: add esp, 00000014h
  loc_0060DEA3: lea ecx, var_24
  loc_0060DEA6: mov edx, [eax]
  loc_0060DEA8: push ecx
  loc_0060DEA9: push eax
  loc_0060DEAA: call [edx+0000008Ch]
  loc_0060DEB0: test eax, eax
  loc_0060DEB2: fnclex
  loc_0060DEB4: jge 0060DECAh
  loc_0060DEB6: mov edx, var_C4
  loc_0060DEBC: push 0000008Ch
  loc_0060DEC1: push 0043B5D0h
  loc_0060DEC6: push edx
  loc_0060DEC7: push eax
  loc_0060DEC8: call edi
  loc_0060DECA: lea ebx, var_28
  loc_0060DECD: mov eax, var_24
  loc_0060DED0: push ebx
  loc_0060DED1: mov edx, 00000008h
  loc_0060DED6: sub esp, 00000010h
  loc_0060DED9: mov var_70, edx
  loc_0060DEDC: mov ebx, esp
  loc_0060DEDE: mov ecx, 0043B848h ; "Section4"
  loc_0060DEE3: mov var_68, ecx
  loc_0060DEE6: mov esi, [eax]
  loc_0060DEE8: mov [ebx], edx
  loc_0060DEEA: mov edx, var_6C
  loc_0060DEED: push eax
  loc_0060DEEE: mov var_AC, eax
  loc_0060DEF4: mov [ebx+00000004h], edx
  loc_0060DEF7: mov [ebx+00000008h], ecx
  loc_0060DEFA: mov ecx, var_64
  loc_0060DEFD: mov [ebx+0000000Ch], ecx
  loc_0060DF00: call [esi+0000001Ch]
  loc_0060DF03: test eax, eax
  loc_0060DF05: fnclex
  loc_0060DF07: jge 0060DF1Ah
  loc_0060DF09: mov edx, var_AC
  loc_0060DF0F: push 0000001Ch
  loc_0060DF11: push 00439DD8h
  loc_0060DF16: push edx
  loc_0060DF17: push eax
  loc_0060DF18: call edi
  loc_0060DF1A: mov eax, var_28
  loc_0060DF1D: lea edx, var_2C
  loc_0060DF20: push edx
  loc_0060DF21: push eax
  loc_0060DF22: mov ecx, [eax]
  loc_0060DF24: mov esi, eax
  loc_0060DF26: call [ecx+0000003Ch]
  loc_0060DF29: test eax, eax
  loc_0060DF2B: fnclex
  loc_0060DF2D: jge 0060DF3Ah
  loc_0060DF2F: push 0000003Ch
  loc_0060DF31: push 00439BF8h
  loc_0060DF36: push esi
  loc_0060DF37: push eax
  loc_0060DF38: call edi
  loc_0060DF3A: lea ebx, var_30
  loc_0060DF3D: mov eax, var_2C
  loc_0060DF40: push ebx
  loc_0060DF41: mov edx, 00000008h
  loc_0060DF46: sub esp, 00000010h
  loc_0060DF49: mov esi, [eax]
  loc_0060DF4B: mov ebx, esp
  loc_0060DF4D: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060DF52: push eax
  loc_0060DF53: mov var_BC, eax
  loc_0060DF59: mov [ebx], edx
  loc_0060DF5B: mov edx, var_7C
  loc_0060DF5E: mov [ebx+00000004h], edx
  loc_0060DF61: mov [ebx+00000008h], ecx
  loc_0060DF64: mov ecx, var_74
  loc_0060DF67: mov [ebx+0000000Ch], ecx
  loc_0060DF6A: call [esi+0000001Ch]
  loc_0060DF6D: test eax, eax
  loc_0060DF6F: fnclex
  loc_0060DF71: jge 0060DF84h
  loc_0060DF73: mov edx, var_BC
  loc_0060DF79: push 0000001Ch
  loc_0060DF7B: push 0043B7D8h
  loc_0060DF80: push edx
  loc_0060DF81: push eax
  loc_0060DF82: call edi
  loc_0060DF84: mov ecx, [006680F8h]
  loc_0060DF8A: mov edx, [006680FCh]
  loc_0060DF90: sub esp, 00000010h
  loc_0060DF93: mov eax, esp
  loc_0060DF95: push 0043B85Ch ; "Caption"
  loc_0060DF9A: mov [eax], ecx
  loc_0060DF9C: mov ecx, [00668100h]
  loc_0060DFA2: mov [eax+00000004h], edx
  loc_0060DFA5: mov edx, [00668104h]
  loc_0060DFAB: mov [eax+00000008h], ecx
  loc_0060DFAE: mov [eax+0000000Ch], edx
  loc_0060DFB1: mov eax, var_30
  loc_0060DFB4: push eax
  loc_0060DFB5: call [00401094h] ; __vbaLateMemSt
  loc_0060DFBB: lea ecx, var_30
  loc_0060DFBE: lea edx, var_2C
  loc_0060DFC1: push ecx
  loc_0060DFC2: lea eax, var_28
  loc_0060DFC5: push edx
  loc_0060DFC6: lea ecx, var_24
  loc_0060DFC9: push eax
  loc_0060DFCA: push ecx
  loc_0060DFCB: push 00000004h
  loc_0060DFCD: call [0040104Ch] ; __vbaFreeObjList
  loc_0060DFD3: mov eax, var_C4
  loc_0060DFD9: add esp, 00000014h
  loc_0060DFDC: lea ecx, var_24
  loc_0060DFDF: mov edx, [eax]
  loc_0060DFE1: push ecx
  loc_0060DFE2: push eax
  loc_0060DFE3: call [edx+0000008Ch]
  loc_0060DFE9: test eax, eax
  loc_0060DFEB: fnclex
  loc_0060DFED: jge 0060E003h
  loc_0060DFEF: mov edx, var_C4
  loc_0060DFF5: push 0000008Ch
  loc_0060DFFA: push 0043B5D0h
  loc_0060DFFF: push edx
  loc_0060E000: push eax
  loc_0060E001: call edi
  loc_0060E003: lea ebx, var_28
  loc_0060E006: mov eax, var_24
  loc_0060E009: push ebx
  loc_0060E00A: mov edx, 00000008h
  loc_0060E00F: sub esp, 00000010h
  loc_0060E012: mov var_70, edx
  loc_0060E015: mov ebx, esp
  loc_0060E017: mov ecx, 0043B848h ; "Section4"
  loc_0060E01C: mov var_68, ecx
  loc_0060E01F: mov esi, [eax]
  loc_0060E021: mov [ebx], edx
  loc_0060E023: mov edx, var_6C
  loc_0060E026: push eax
  loc_0060E027: mov var_AC, eax
  loc_0060E02D: mov [ebx+00000004h], edx
  loc_0060E030: mov [ebx+00000008h], ecx
  loc_0060E033: mov ecx, var_64
  loc_0060E036: mov [ebx+0000000Ch], ecx
  loc_0060E039: call [esi+0000001Ch]
  loc_0060E03C: test eax, eax
  loc_0060E03E: fnclex
  loc_0060E040: jge 0060E053h
  loc_0060E042: mov edx, var_AC
  loc_0060E048: push 0000001Ch
  loc_0060E04A: push 00439DD8h
  loc_0060E04F: push edx
  loc_0060E050: push eax
  loc_0060E051: call edi
  loc_0060E053: mov eax, var_28
  loc_0060E056: lea edx, var_2C
  loc_0060E059: push edx
  loc_0060E05A: push eax
  loc_0060E05B: mov ecx, [eax]
  loc_0060E05D: mov esi, eax
  loc_0060E05F: call [ecx+0000003Ch]
  loc_0060E062: test eax, eax
  loc_0060E064: fnclex
  loc_0060E066: jge 0060E073h
  loc_0060E068: push 0000003Ch
  loc_0060E06A: push 00439BF8h
  loc_0060E06F: push esi
  loc_0060E070: push eax
  loc_0060E071: call edi
  loc_0060E073: lea ebx, var_30
  loc_0060E076: mov eax, var_2C
  loc_0060E079: push ebx
  loc_0060E07A: mov edx, 00000008h
  loc_0060E07F: sub esp, 00000010h
  loc_0060E082: mov esi, [eax]
  loc_0060E084: mov ebx, esp
  loc_0060E086: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060E08B: push eax
  loc_0060E08C: mov var_BC, eax
  loc_0060E092: mov [ebx], edx
  loc_0060E094: mov edx, var_7C
  loc_0060E097: mov [ebx+00000004h], edx
  loc_0060E09A: mov [ebx+00000008h], ecx
  loc_0060E09D: mov ecx, var_74
  loc_0060E0A0: mov [ebx+0000000Ch], ecx
  loc_0060E0A3: call [esi+0000001Ch]
  loc_0060E0A6: test eax, eax
  loc_0060E0A8: fnclex
  loc_0060E0AA: jge 0060E0BDh
  loc_0060E0AC: mov edx, var_BC
  loc_0060E0B2: push 0000001Ch
  loc_0060E0B4: push 0043B7D8h
  loc_0060E0B9: push edx
  loc_0060E0BA: push eax
  loc_0060E0BB: call edi
  loc_0060E0BD: mov ecx, [00668108h]
  loc_0060E0C3: mov edx, [0066810Ch]
  loc_0060E0C9: sub esp, 00000010h
  loc_0060E0CC: mov eax, esp
  loc_0060E0CE: push 0043B85Ch ; "Caption"
  loc_0060E0D3: mov [eax], ecx
  loc_0060E0D5: mov ecx, [00668110h]
  loc_0060E0DB: mov [eax+00000004h], edx
  loc_0060E0DE: mov edx, [00668114h]
  loc_0060E0E4: mov [eax+00000008h], ecx
  loc_0060E0E7: mov [eax+0000000Ch], edx
  loc_0060E0EA: mov eax, var_30
  loc_0060E0ED: push eax
  loc_0060E0EE: call [00401094h] ; __vbaLateMemSt
  loc_0060E0F4: lea ecx, var_30
  loc_0060E0F7: lea edx, var_2C
  loc_0060E0FA: push ecx
  loc_0060E0FB: lea eax, var_28
  loc_0060E0FE: push edx
  loc_0060E0FF: lea ecx, var_24
  loc_0060E102: push eax
  loc_0060E103: push ecx
  loc_0060E104: push 00000004h
  loc_0060E106: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E10C: mov eax, var_C4
  loc_0060E112: add esp, 00000014h
  loc_0060E115: mov edx, [eax]
  loc_0060E117: push 0000000Ah
  loc_0060E119: push eax
  loc_0060E11A: call [edx+00000048h]
  loc_0060E11D: test eax, eax
  loc_0060E11F: fnclex
  loc_0060E121: jge 0060E134h
  loc_0060E123: mov ecx, var_C4
  loc_0060E129: push 00000048h
  loc_0060E12B: push 0043B5D0h
  loc_0060E130: push ecx
  loc_0060E131: push eax
  loc_0060E132: call edi
  loc_0060E134: mov eax, var_C4
  loc_0060E13A: push 0000000Ah
  loc_0060E13C: push eax
  loc_0060E13D: mov edx, [eax]
  loc_0060E13F: call [edx+00000058h]
  loc_0060E142: test eax, eax
  loc_0060E144: fnclex
  loc_0060E146: jge 0060E159h
  loc_0060E148: mov ecx, var_C4
  loc_0060E14E: push 00000058h
  loc_0060E150: push 0043B5D0h
  loc_0060E155: push ecx
  loc_0060E156: push eax
  loc_0060E157: call edi
  loc_0060E159: mov edx, var_C4
  loc_0060E15F: push 00000000h
  loc_0060E161: push 80011003h
  loc_0060E166: push edx
  loc_0060E167: call [0040102Ch] ; __vbaLateIdCall
  loc_0060E16D: add esp, 0000000Ch
  loc_0060E170: lea eax, var_C4
  loc_0060E176: push 00000000h
  loc_0060E178: push eax
  loc_0060E179: call [004010C8h] ; __vbaObjSetAddref
  loc_0060E17F: push 0060E1E9h
  loc_0060E184: jmp 0060E1BFh
  loc_0060E186: lea ecx, var_20
  loc_0060E189: call [0040129Ch] ; __vbaFreeStr
  loc_0060E18F: lea ecx, var_30
  loc_0060E192: lea edx, var_2C
  loc_0060E195: push ecx
  loc_0060E196: lea eax, var_28
  loc_0060E199: push edx
  loc_0060E19A: lea ecx, var_24
  loc_0060E19D: push eax
  loc_0060E19E: push ecx
  loc_0060E19F: push 00000004h
  loc_0060E1A1: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E1A7: lea edx, var_60
  loc_0060E1AA: lea eax, var_50
  loc_0060E1AD: push edx
  loc_0060E1AE: lea ecx, var_40
  loc_0060E1B1: push eax
  loc_0060E1B2: push ecx
  loc_0060E1B3: push 00000003h
  loc_0060E1B5: call [0040103Ch] ; __vbaFreeVarList
  loc_0060E1BB: add esp, 00000024h
  loc_0060E1BE: ret
  loc_0060E1BF: lea edx, var_C8
  loc_0060E1C5: lea eax, var_C4
  loc_0060E1CB: push edx
  loc_0060E1CC: push eax
  loc_0060E1CD: push 00000002h
  loc_0060E1CF: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E1D5: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060E1DB: add esp, 0000000Ch
  loc_0060E1DE: lea ecx, var_18
  loc_0060E1E1: call __vbaFreeStr
  loc_0060E1E3: lea ecx, var_1C
  loc_0060E1E6: call __vbaFreeStr
  loc_0060E1E8: ret
  loc_0060E1E9: mov ecx, var_10
  loc_0060E1EC: pop edi
  loc_0060E1ED: pop esi
  loc_0060E1EE: mov fs:[00000000h], ecx
  loc_0060E1F5: pop ebx
  loc_0060E1F6: mov esp, ebp
  loc_0060E1F8: pop ebp
  loc_0060E1F9: ret
End Sub

Private Sub Proc_34_5_60E200
  loc_0060E200: push ebp
  loc_0060E201: mov ebp, esp
  loc_0060E203: sub esp, 00000008h
  loc_0060E206: push 00405696h ; __vbaExceptHandler
  loc_0060E20B: mov eax, fs:[00000000h]
  loc_0060E211: push eax
  loc_0060E212: mov fs:[00000000h], esp
  loc_0060E219: sub esp, 000000D0h
  loc_0060E21F: push ebx
  loc_0060E220: push esi
  loc_0060E221: push edi
  loc_0060E222: mov var_8, esp
  loc_0060E225: mov var_4, 004044F0h
  loc_0060E22C: xor ebx, ebx
  loc_0060E22E: mov var_18, ebx
  loc_0060E231: mov var_1C, ebx
  loc_0060E234: mov var_20, ebx
  loc_0060E237: mov var_24, ebx
  loc_0060E23A: mov var_28, ebx
  loc_0060E23D: mov var_2C, ebx
  loc_0060E240: mov var_30, ebx
  loc_0060E243: mov var_40, ebx
  loc_0060E246: mov var_50, ebx
  loc_0060E249: mov var_60, ebx
  loc_0060E24C: mov var_70, ebx
  loc_0060E24F: mov var_80, ebx
  loc_0060E252: mov var_90, ebx
  loc_0060E258: mov var_A4, ebx
  loc_0060E25E: mov var_C4, ebx
  loc_0060E264: mov var_C8, ebx
  loc_0060E26A: call 0055B320h
  loc_0060E26F: mov edx, 0042DFECh
  loc_0060E274: lea ecx, var_18
  loc_0060E277: call [004011FCh] ; __vbaStrCopy
  loc_0060E27D: push 0043BC58h ; " SELECT kode_barang, nama_barang, satuan, harga_jual, diskon, stok"
  loc_0060E282: push 0043BD38h ; " FROM brg_barang ORDER BY kode_barang, kode_produk"
  loc_0060E287: call [00401070h] ; __vbaStrCat
  loc_0060E28D: mov edx, eax
  loc_0060E28F: lea ecx, var_18
  loc_0060E292: call [0040126Ch] ; __vbaStrMove
  loc_0060E298: push 0042DF30h
  loc_0060E29D: call [00401168h] ; __vbaNew
  loc_0060E2A3: push eax
  loc_0060E2A4: push 00668090h
  loc_0060E2A9: call [004010B8h] ; __vbaObjSet
  loc_0060E2AF: cmp [0066803Ch], ebx
  loc_0060E2B5: jnz 0060E2C7h
  loc_0060E2B7: push 0066803Ch
  loc_0060E2BC: push 0042DEFCh
  loc_0060E2C1: call [004011E8h] ; __vbaNew2
  loc_0060E2C7: mov esi, [0066803Ch]
  loc_0060E2CD: lea ecx, var_40
  loc_0060E2D0: mov var_38, 80020004h
  loc_0060E2D7: mov var_40, 0000000Ah
  loc_0060E2DE: call [0040123Ch] ; __vbaFreeVarg
  loc_0060E2E4: mov eax, [esi]
  loc_0060E2E6: lea ecx, var_24
  loc_0060E2E9: push ecx
  loc_0060E2EA: mov ecx, var_18
  loc_0060E2ED: lea edx, var_40
  loc_0060E2F0: push 00000001h
  loc_0060E2F2: push edx
  loc_0060E2F3: push ecx
  loc_0060E2F4: push esi
  loc_0060E2F5: call [eax+00000040h]
  loc_0060E2F8: cmp eax, ebx
  loc_0060E2FA: fnclex
  loc_0060E2FC: jge 0060E311h
  loc_0060E2FE: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060E304: push 00000040h
  loc_0060E306: push 0042E1B0h
  loc_0060E30B: push esi
  loc_0060E30C: push eax
  loc_0060E30D: call edi
  loc_0060E30F: jmp 0060E317h
  loc_0060E311: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060E317: mov edx, var_24
  loc_0060E31A: mov esi, [00401268h] ; __vbaCastObj
  loc_0060E320: push 0042DF20h
  loc_0060E325: push edx
  loc_0060E326: call __vbaCastObj
  loc_0060E328: push eax
  loc_0060E329: push 00668090h
  loc_0060E32E: call [004010B8h] ; __vbaObjSet
  loc_0060E334: lea ecx, var_24
  loc_0060E337: call [004012A0h] ; __vbaFreeObj
  loc_0060E33D: lea ecx, var_40
  loc_0060E340: call [00401020h] ; __vbaFreeVar
  loc_0060E346: cmp [00668558h], ebx
  loc_0060E34C: jnz 0060E35Eh
  loc_0060E34E: push 00668558h
  loc_0060E353: push 004080B8h
  loc_0060E358: call [004011E8h] ; __vbaNew2
  loc_0060E35E: mov eax, [00668558h]
  loc_0060E363: lea ecx, var_C4
  loc_0060E369: push eax
  loc_0060E36A: push ecx
  loc_0060E36B: call [004010C8h] ; __vbaObjSetAddref
  loc_0060E371: mov edx, var_C4
  loc_0060E377: push 0043B7B8h
  loc_0060E37C: push ebx
  loc_0060E37D: mov edx, [edx]
  loc_0060E37F: mov var_E0, edx
  loc_0060E385: call __vbaCastObj
  loc_0060E387: push eax
  loc_0060E388: lea eax, var_24
  loc_0060E38B: push eax
  loc_0060E38C: call [004010B8h] ; __vbaObjSet
  loc_0060E392: mov ecx, var_C4
  loc_0060E398: mov edx, var_E0
  loc_0060E39E: push eax
  loc_0060E39F: push ecx
  loc_0060E3A0: call [edx+00000028h]
  loc_0060E3A3: cmp eax, ebx
  loc_0060E3A5: fnclex
  loc_0060E3A7: jge 0060E3BAh
  loc_0060E3A9: mov ecx, var_C4
  loc_0060E3AF: push 00000028h
  loc_0060E3B1: push 0043B5D0h
  loc_0060E3B6: push ecx
  loc_0060E3B7: push eax
  loc_0060E3B8: call edi
  loc_0060E3BA: lea ecx, var_24
  loc_0060E3BD: call [004012A0h] ; __vbaFreeObj
  loc_0060E3C3: mov eax, var_C4
  loc_0060E3C9: push 0042DFECh
  loc_0060E3CE: push eax
  loc_0060E3CF: mov edx, [eax]
  loc_0060E3D1: call [edx+00000030h]
  loc_0060E3D4: cmp eax, ebx
  loc_0060E3D6: fnclex
  loc_0060E3D8: jge 0060E3EBh
  loc_0060E3DA: mov ecx, var_C4
  loc_0060E3E0: push 00000030h
  loc_0060E3E2: push 0043B5D0h
  loc_0060E3E7: push ecx
  loc_0060E3E8: push eax
  loc_0060E3E9: call edi
  loc_0060E3EB: mov eax, [00668090h]
  loc_0060E3F0: lea ecx, var_24
  loc_0060E3F3: push ecx
  loc_0060E3F4: push eax
  loc_0060E3F5: mov edx, [eax]
  loc_0060E3F7: call [edx+00000114h]
  loc_0060E3FD: cmp eax, ebx
  loc_0060E3FF: fnclex
  loc_0060E401: jge 0060E417h
  loc_0060E403: mov edx, [00668090h]
  loc_0060E409: push 00000114h
  loc_0060E40E: push 0043B7C8h
  loc_0060E413: push edx
  loc_0060E414: push eax
  loc_0060E415: call edi
  loc_0060E417: mov eax, var_C4
  loc_0060E41D: mov ecx, var_24
  loc_0060E420: push 0043B7B8h
  loc_0060E425: push ecx
  loc_0060E426: mov esi, [eax]
  loc_0060E428: call [00401268h] ; __vbaCastObj
  loc_0060E42E: lea edx, var_28
  loc_0060E431: push eax
  loc_0060E432: push edx
  loc_0060E433: call [004010B8h] ; __vbaObjSet
  loc_0060E439: push eax
  loc_0060E43A: mov eax, var_C4
  loc_0060E440: push eax
  loc_0060E441: call [esi+00000028h]
  loc_0060E444: cmp eax, ebx
  loc_0060E446: fnclex
  loc_0060E448: jge 0060E45Bh
  loc_0060E44A: mov ecx, var_C4
  loc_0060E450: push 00000028h
  loc_0060E452: push 0043B5D0h
  loc_0060E457: push ecx
  loc_0060E458: push eax
  loc_0060E459: call edi
  loc_0060E45B: lea edx, var_28
  loc_0060E45E: lea eax, var_24
  loc_0060E461: push edx
  loc_0060E462: push eax
  loc_0060E463: push 00000002h
  loc_0060E465: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E46B: mov eax, var_C4
  loc_0060E471: add esp, 0000000Ch
  loc_0060E474: lea edx, var_24
  loc_0060E477: mov ecx, [eax]
  loc_0060E479: push edx
  loc_0060E47A: push eax
  loc_0060E47B: call [ecx+0000008Ch]
  loc_0060E481: cmp eax, ebx
  loc_0060E483: fnclex
  loc_0060E485: jge 0060E49Bh
  loc_0060E487: mov ecx, var_C4
  loc_0060E48D: push 0000008Ch
  loc_0060E492: push 0043B5D0h
  loc_0060E497: push ecx
  loc_0060E498: push eax
  loc_0060E499: call edi
  loc_0060E49B: lea esi, var_28
  loc_0060E49E: mov eax, var_24
  loc_0060E4A1: push esi
  loc_0060E4A2: mov ecx, 00000008h
  loc_0060E4A7: sub esp, 00000010h
  loc_0060E4AA: mov var_70, ecx
  loc_0060E4AD: mov esi, esp
  loc_0060E4AF: mov var_68, 00439FB0h ; "Section1"
  loc_0060E4B6: mov edx, [eax]
  loc_0060E4B8: push eax
  loc_0060E4B9: mov [esi], ecx
  loc_0060E4BB: mov ecx, var_6C
  loc_0060E4BE: mov var_AC, eax
  loc_0060E4C4: mov [esi+00000004h], ecx
  loc_0060E4C7: mov ecx, var_68
  loc_0060E4CA: mov [esi+00000008h], ecx
  loc_0060E4CD: mov ecx, var_64
  loc_0060E4D0: mov [esi+0000000Ch], ecx
  loc_0060E4D3: call [edx+0000001Ch]
  loc_0060E4D6: cmp eax, ebx
  loc_0060E4D8: fnclex
  loc_0060E4DA: jge 0060E4EDh
  loc_0060E4DC: mov edx, var_AC
  loc_0060E4E2: push 0000001Ch
  loc_0060E4E4: push 00439DD8h
  loc_0060E4E9: push edx
  loc_0060E4EA: push eax
  loc_0060E4EB: call edi
  loc_0060E4ED: mov eax, var_28
  loc_0060E4F0: lea edx, var_2C
  loc_0060E4F3: push edx
  loc_0060E4F4: push eax
  loc_0060E4F5: mov ecx, [eax]
  loc_0060E4F7: mov esi, eax
  loc_0060E4F9: call [ecx+0000003Ch]
  loc_0060E4FC: cmp eax, ebx
  loc_0060E4FE: fnclex
  loc_0060E500: jge 0060E50Dh
  loc_0060E502: push 0000003Ch
  loc_0060E504: push 00439BF8h
  loc_0060E509: push esi
  loc_0060E50A: push eax
  loc_0060E50B: call edi
  loc_0060E50D: mov eax, var_2C
  loc_0060E510: mov var_2C, ebx
  loc_0060E513: push eax
  loc_0060E514: lea eax, var_C8
  loc_0060E51A: push eax
  loc_0060E51B: call [004010B8h] ; __vbaObjSet
  loc_0060E521: lea ecx, var_28
  loc_0060E524: lea edx, var_24
  loc_0060E527: push ecx
  loc_0060E528: push edx
  loc_0060E529: push 00000002h
  loc_0060E52B: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E531: mov eax, var_C8
  loc_0060E537: add esp, 0000000Ch
  loc_0060E53A: lea edx, var_A4
  loc_0060E540: mov ecx, [eax]
  loc_0060E542: push edx
  loc_0060E543: push eax
  loc_0060E544: call [ecx+00000020h]
  loc_0060E547: cmp eax, ebx
  loc_0060E549: fnclex
  loc_0060E54B: jge 0060E55Eh
  loc_0060E54D: mov ecx, var_C8
  loc_0060E553: push 00000020h
  loc_0060E555: push 0043B7D8h
  loc_0060E55A: push ecx
  loc_0060E55B: push eax
  loc_0060E55C: call edi
  loc_0060E55E: mov ecx, var_A4
  loc_0060E564: call [00401138h] ; __vbaI2I4
  loc_0060E56A: mov var_D0, eax
  loc_0060E570: mov ebx, 00000001h
  loc_0060E575: cmp bx, var_D0
  loc_0060E57C: mov var_14, ebx
  loc_0060E57F: jg 0060E7F0h
  loc_0060E585: lea esi, var_24
  loc_0060E588: mov ecx, var_C8
  loc_0060E58E: push esi
  loc_0060E58F: mov eax, 00000002h
  loc_0060E594: sub esp, 00000010h
  loc_0060E597: mov var_70, eax
  loc_0060E59A: mov esi, esp
  loc_0060E59C: mov var_68, bx
  loc_0060E5A0: mov edx, [ecx]
  loc_0060E5A2: push ecx
  loc_0060E5A3: mov [esi], eax
  loc_0060E5A5: mov eax, var_6C
  loc_0060E5A8: mov [esi+00000004h], eax
  loc_0060E5AB: mov eax, var_68
  loc_0060E5AE: mov [esi+00000008h], eax
  loc_0060E5B1: mov eax, var_64
  loc_0060E5B4: mov [esi+0000000Ch], eax
  loc_0060E5B7: call [edx+0000001Ch]
  loc_0060E5BA: test eax, eax
  loc_0060E5BC: fnclex
  loc_0060E5BE: jge 0060E5D1h
  loc_0060E5C0: mov ecx, var_C8
  loc_0060E5C6: push 0000001Ch
  loc_0060E5C8: push 0043B7D8h
  loc_0060E5CD: push ecx
  loc_0060E5CE: push eax
  loc_0060E5CF: call edi
  loc_0060E5D1: mov edx, var_24
  loc_0060E5D4: push 0043B7E8h
  loc_0060E5D9: push edx
  loc_0060E5DA: call [004011CCh] ; __vbaCheckType
  loc_0060E5E0: lea ecx, var_24
  loc_0060E5E3: mov si, ax
  loc_0060E5E6: call [004012A0h] ; __vbaFreeObj
  loc_0060E5EC: test si, si
  loc_0060E5EF: Unknown_795()
  loc_0060E5F5: mov var_68, bx
  loc_0060E5F9: lea ebx, var_24
  loc_0060E5FC: push ebx
  loc_0060E5FD: mov ecx, var_C8
  loc_0060E603: sub esp, 00000010h
  loc_0060E606: mov eax, 00000002h
  loc_0060E60B: mov ebx, esp
  loc_0060E60D: mov var_70, eax
  loc_0060E610: mov edx, [ecx]
  loc_0060E612: mov esi, 0042DFECh
  loc_0060E617: mov [ebx], eax
  loc_0060E619: mov eax, var_6C
  loc_0060E61C: push ecx
  loc_0060E61D: mov var_78, esi
  loc_0060E620: mov [ebx+00000004h], eax
  loc_0060E623: mov eax, var_68
  loc_0060E626: mov edi, 00000008h
  loc_0060E62B: mov [ebx+00000008h], eax
  loc_0060E62E: mov eax, var_64
  loc_0060E631: mov [ebx+0000000Ch], eax
  loc_0060E634: call [edx+0000001Ch]
  loc_0060E637: test eax, eax
  loc_0060E639: fnclex
  loc_0060E63B: jge 0060E652h
  loc_0060E63D: mov ecx, var_C8
  loc_0060E643: push 0000001Ch
  loc_0060E645: push 0043B7D8h
  loc_0060E64A: push ecx
  loc_0060E64B: push eax
  loc_0060E64C: call [00401080h] ; __vbaHresultCheckObj
  loc_0060E652: mov eax, var_7C
  loc_0060E655: sub esp, 00000010h
  loc_0060E658: mov edx, esp
  loc_0060E65A: mov ecx, var_74
  loc_0060E65D: push 0043B7F8h ; "DataMember"
  loc_0060E662: mov [edx], edi
  loc_0060E664: mov [edx+00000004h], eax
  loc_0060E667: mov [edx+00000008h], esi
  loc_0060E66A: mov [edx+0000000Ch], ecx
  loc_0060E66D: mov edx, var_24
  loc_0060E670: push edx
  loc_0060E671: call [00401094h] ; __vbaLateMemSt
  loc_0060E677: lea ecx, var_24
  loc_0060E67A: call [004012A0h] ; __vbaFreeObj
  loc_0060E680: mov eax, [00668090h]
  loc_0060E685: lea edx, var_24
  loc_0060E688: push edx
  loc_0060E689: push eax
  loc_0060E68A: mov ecx, [eax]
  loc_0060E68C: call [ecx+00000054h]
  loc_0060E68F: test eax, eax
  loc_0060E691: fnclex
  loc_0060E693: jge 0060E6AAh
  loc_0060E695: mov ecx, [00668090h]
  loc_0060E69B: push 00000054h
  loc_0060E69D: push 0042DF44h
  loc_0060E6A2: push ecx
  loc_0060E6A3: push eax
  loc_0060E6A4: call [00401080h] ; __vbaHresultCheckObj
  loc_0060E6AA: mov ebx, var_14
  loc_0060E6AD: lea edi, var_28
  loc_0060E6B0: mov dx, bx
  loc_0060E6B3: push edi
  loc_0060E6B4: sub dx, 0001h
  loc_0060E6B8: mov eax, var_24
  loc_0060E6BB: jo 0060EF7Eh
  loc_0060E6C1: sub esp, 00000010h
  loc_0060E6C4: mov ecx, 00000002h
  loc_0060E6C9: mov edi, esp
  loc_0060E6CB: mov var_70, ecx
  loc_0060E6CE: mov var_68, dx
  loc_0060E6D2: mov edx, [eax]
  loc_0060E6D4: mov [edi], ecx
  loc_0060E6D6: mov ecx, var_6C
  loc_0060E6D9: push eax
  loc_0060E6DA: mov esi, eax
  loc_0060E6DC: mov [edi+00000004h], ecx
  loc_0060E6DF: mov ecx, var_68
  loc_0060E6E2: mov [edi+00000008h], ecx
  loc_0060E6E5: mov ecx, var_64
  loc_0060E6E8: mov [edi+0000000Ch], ecx
  loc_0060E6EB: call [edx+00000028h]
  loc_0060E6EE: test eax, eax
  loc_0060E6F0: fnclex
  loc_0060E6F2: jge 0060E707h
  loc_0060E6F4: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060E6FA: push 00000028h
  loc_0060E6FC: push 0042DFACh
  loc_0060E701: push esi
  loc_0060E702: push eax
  loc_0060E703: call edi
  loc_0060E705: jmp 0060E70Dh
  loc_0060E707: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060E70D: mov eax, var_28
  loc_0060E710: lea ecx, var_20
  loc_0060E713: push ecx
  loc_0060E714: push eax
  loc_0060E715: mov edx, [eax]
  loc_0060E717: mov esi, eax
  loc_0060E719: call [edx+0000002Ch]
  loc_0060E71C: test eax, eax
  loc_0060E71E: fnclex
  loc_0060E720: jge 0060E72Dh
  loc_0060E722: push 0000002Ch
  loc_0060E724: push 0042DFBCh
  loc_0060E729: push esi
  loc_0060E72A: push eax
  loc_0060E72B: call edi
  loc_0060E72D: mov eax, var_20
  loc_0060E730: lea esi, var_2C
  loc_0060E733: push esi
  loc_0060E734: mov ecx, var_C8
  loc_0060E73A: sub esp, 00000010h
  loc_0060E73D: mov var_38, eax
  loc_0060E740: mov esi, esp
  loc_0060E742: mov eax, 00000002h
  loc_0060E747: mov var_78, bx
  loc_0060E74B: mov var_20, 00000000h
  loc_0060E752: mov [esi], eax
  loc_0060E754: mov eax, var_7C
  loc_0060E757: mov var_40, 00000008h
  loc_0060E75E: mov edx, [ecx]
  loc_0060E760: mov [esi+00000004h], eax
  loc_0060E763: mov eax, var_78
  loc_0060E766: push ecx
  loc_0060E767: mov [esi+00000008h], eax
  loc_0060E76A: mov eax, var_74
  loc_0060E76D: mov [esi+0000000Ch], eax
  loc_0060E770: call [edx+0000001Ch]
  loc_0060E773: test eax, eax
  loc_0060E775: fnclex
  loc_0060E777: jge 0060E78Ah
  loc_0060E779: mov ecx, var_C8
  loc_0060E77F: push 0000001Ch
  loc_0060E781: push 0043B7D8h
  loc_0060E786: push ecx
  loc_0060E787: push eax
  loc_0060E788: call edi
  loc_0060E78A: mov eax, var_40
  loc_0060E78D: mov ecx, var_3C
  loc_0060E790: sub esp, 00000010h
  loc_0060E793: mov edx, esp
  loc_0060E795: push 0043B810h ; "DataField"
  loc_0060E79A: mov [edx], eax
  loc_0060E79C: mov eax, var_38
  loc_0060E79F: mov [edx+00000004h], ecx
  loc_0060E7A2: mov ecx, var_34
  loc_0060E7A5: mov [edx+00000008h], eax
  loc_0060E7A8: mov [edx+0000000Ch], ecx
  loc_0060E7AB: mov edx, var_2C
  loc_0060E7AE: push edx
  loc_0060E7AF: call [00401094h] ; __vbaLateMemSt
  loc_0060E7B5: lea eax, var_2C
  loc_0060E7B8: lea ecx, var_28
  loc_0060E7BB: push eax
  loc_0060E7BC: lea edx, var_24
  loc_0060E7BF: push ecx
  loc_0060E7C0: push edx
  loc_0060E7C1: push 00000003h
  loc_0060E7C3: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E7C9: add esp, 00000010h
  loc_0060E7CC: lea ecx, var_40
  loc_0060E7CF: call [00401020h] ; __vbaFreeVar
  loc_0060E7D5: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060E7DB: mov eax, 00000001h
  loc_0060E7E0: add ax, bx
  loc_0060E7E3: jo 0060EF7Eh
  loc_0060E7E9: mov ebx, eax
  loc_0060E7EB: jmp 0060E575h
  loc_0060E7F0: lea eax, var_C8
  loc_0060E7F6: push 00000000h
  loc_0060E7F8: push eax
  loc_0060E7F9: call [004010C8h] ; __vbaObjSetAddref
  loc_0060E7FF: lea ecx, var_40
  loc_0060E802: push ecx
  loc_0060E803: call [00401220h] ; rtcGetDateVar
  loc_0060E809: lea edx, var_70
  loc_0060E80C: lea ecx, var_50
  loc_0060E80F: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060E816: mov var_70, 00000008h
  loc_0060E81D: call [00401238h] ; __vbaVarDup
  loc_0060E823: push 00000001h
  loc_0060E825: lea edx, var_50
  loc_0060E828: push 00000001h
  loc_0060E82A: lea eax, var_40
  loc_0060E82D: push edx
  loc_0060E82E: lea ecx, var_60
  loc_0060E831: push eax
  loc_0060E832: push ecx
  loc_0060E833: call [00401078h] ; rtcVarFromFormatVar
  loc_0060E839: mov eax, var_C4
  loc_0060E83F: lea ecx, var_24
  loc_0060E842: push ecx
  loc_0060E843: push eax
  loc_0060E844: mov edx, [eax]
  loc_0060E846: call [edx+0000008Ch]
  loc_0060E84C: test eax, eax
  loc_0060E84E: fnclex
  loc_0060E850: jge 0060E866h
  loc_0060E852: mov edx, var_C4
  loc_0060E858: push 0000008Ch
  loc_0060E85D: push 0043B5D0h
  loc_0060E862: push edx
  loc_0060E863: push eax
  loc_0060E864: call edi
  loc_0060E866: lea ebx, var_28
  loc_0060E869: mov eax, var_24
  loc_0060E86C: push ebx
  loc_0060E86D: mov edx, 00000008h
  loc_0060E872: sub esp, 00000010h
  loc_0060E875: mov esi, [eax]
  loc_0060E877: mov ebx, esp
  loc_0060E879: mov ecx, 0043B848h ; "Section4"
  loc_0060E87E: push eax
  loc_0060E87F: mov var_AC, eax
  loc_0060E885: mov [ebx], edx
  loc_0060E887: mov edx, var_7C
  loc_0060E88A: mov [ebx+00000004h], edx
  loc_0060E88D: mov [ebx+00000008h], ecx
  loc_0060E890: mov ecx, var_74
  loc_0060E893: mov [ebx+0000000Ch], ecx
  loc_0060E896: call [esi+0000001Ch]
  loc_0060E899: test eax, eax
  loc_0060E89B: fnclex
  loc_0060E89D: jge 0060E8B0h
  loc_0060E89F: mov edx, var_AC
  loc_0060E8A5: push 0000001Ch
  loc_0060E8A7: push 00439DD8h
  loc_0060E8AC: push edx
  loc_0060E8AD: push eax
  loc_0060E8AE: call edi
  loc_0060E8B0: mov eax, var_28
  loc_0060E8B3: lea edx, var_2C
  loc_0060E8B6: push edx
  loc_0060E8B7: push eax
  loc_0060E8B8: mov ecx, [eax]
  loc_0060E8BA: mov esi, eax
  loc_0060E8BC: call [ecx+0000003Ch]
  loc_0060E8BF: test eax, eax
  loc_0060E8C1: fnclex
  loc_0060E8C3: jge 0060E8D0h
  loc_0060E8C5: push 0000003Ch
  loc_0060E8C7: push 00439BF8h
  loc_0060E8CC: push esi
  loc_0060E8CD: push eax
  loc_0060E8CE: call edi
  loc_0060E8D0: lea ebx, var_30
  loc_0060E8D3: mov eax, var_2C
  loc_0060E8D6: push ebx
  loc_0060E8D7: mov edx, 00000008h
  loc_0060E8DC: sub esp, 00000010h
  loc_0060E8DF: mov esi, [eax]
  loc_0060E8E1: mov ebx, esp
  loc_0060E8E3: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060E8E8: push eax
  loc_0060E8E9: mov var_BC, eax
  loc_0060E8EF: mov [ebx], edx
  loc_0060E8F1: mov edx, var_8C
  loc_0060E8F7: mov [ebx+00000004h], edx
  loc_0060E8FA: mov [ebx+00000008h], ecx
  loc_0060E8FD: mov ecx, var_84
  loc_0060E903: mov [ebx+0000000Ch], ecx
  loc_0060E906: call [esi+0000001Ch]
  loc_0060E909: test eax, eax
  loc_0060E90B: fnclex
  loc_0060E90D: jge 0060E920h
  loc_0060E90F: mov edx, var_BC
  loc_0060E915: push 0000001Ch
  loc_0060E917: push 0043B7D8h
  loc_0060E91C: push edx
  loc_0060E91D: push eax
  loc_0060E91E: call edi
  loc_0060E920: mov ecx, var_60
  loc_0060E923: mov edx, var_5C
  loc_0060E926: sub esp, 00000010h
  loc_0060E929: mov eax, esp
  loc_0060E92B: push 0043B85Ch ; "Caption"
  loc_0060E930: mov [eax], ecx
  loc_0060E932: mov ecx, var_58
  loc_0060E935: mov [eax+00000004h], edx
  loc_0060E938: mov edx, var_54
  loc_0060E93B: mov [eax+00000008h], ecx
  loc_0060E93E: mov [eax+0000000Ch], edx
  loc_0060E941: mov eax, var_30
  loc_0060E944: push eax
  loc_0060E945: call [00401094h] ; __vbaLateMemSt
  loc_0060E94B: lea ecx, var_30
  loc_0060E94E: lea edx, var_2C
  loc_0060E951: push ecx
  loc_0060E952: lea eax, var_28
  loc_0060E955: push edx
  loc_0060E956: lea ecx, var_24
  loc_0060E959: push eax
  loc_0060E95A: push ecx
  loc_0060E95B: push 00000004h
  loc_0060E95D: call [0040104Ch] ; __vbaFreeObjList
  loc_0060E963: lea edx, var_60
  loc_0060E966: lea eax, var_50
  loc_0060E969: push edx
  loc_0060E96A: lea ecx, var_40
  loc_0060E96D: push eax
  loc_0060E96E: push ecx
  loc_0060E96F: push 00000003h
  loc_0060E971: call [0040103Ch] ; __vbaFreeVarList
  loc_0060E977: mov eax, var_C4
  loc_0060E97D: add esp, 00000024h
  loc_0060E980: lea ecx, var_24
  loc_0060E983: mov edx, [eax]
  loc_0060E985: push ecx
  loc_0060E986: push eax
  loc_0060E987: call [edx+0000008Ch]
  loc_0060E98D: test eax, eax
  loc_0060E98F: fnclex
  loc_0060E991: jge 0060E9A7h
  loc_0060E993: mov edx, var_C4
  loc_0060E999: push 0000008Ch
  loc_0060E99E: push 0043B5D0h
  loc_0060E9A3: push edx
  loc_0060E9A4: push eax
  loc_0060E9A5: call edi
  loc_0060E9A7: lea ebx, var_28
  loc_0060E9AA: mov eax, var_24
  loc_0060E9AD: push ebx
  loc_0060E9AE: mov edx, 00000008h
  loc_0060E9B3: sub esp, 00000010h
  loc_0060E9B6: mov var_70, edx
  loc_0060E9B9: mov ebx, esp
  loc_0060E9BB: mov ecx, 0043B848h ; "Section4"
  loc_0060E9C0: mov var_68, ecx
  loc_0060E9C3: mov esi, [eax]
  loc_0060E9C5: mov [ebx], edx
  loc_0060E9C7: mov edx, var_6C
  loc_0060E9CA: push eax
  loc_0060E9CB: mov var_AC, eax
  loc_0060E9D1: mov [ebx+00000004h], edx
  loc_0060E9D4: mov [ebx+00000008h], ecx
  loc_0060E9D7: mov ecx, var_64
  loc_0060E9DA: mov [ebx+0000000Ch], ecx
  loc_0060E9DD: call [esi+0000001Ch]
  loc_0060E9E0: test eax, eax
  loc_0060E9E2: fnclex
  loc_0060E9E4: jge 0060E9F7h
  loc_0060E9E6: mov edx, var_AC
  loc_0060E9EC: push 0000001Ch
  loc_0060E9EE: push 00439DD8h
  loc_0060E9F3: push edx
  loc_0060E9F4: push eax
  loc_0060E9F5: call edi
  loc_0060E9F7: mov eax, var_28
  loc_0060E9FA: lea edx, var_2C
  loc_0060E9FD: push edx
  loc_0060E9FE: push eax
  loc_0060E9FF: mov ecx, [eax]
  loc_0060EA01: mov esi, eax
  loc_0060EA03: call [ecx+0000003Ch]
  loc_0060EA06: test eax, eax
  loc_0060EA08: fnclex
  loc_0060EA0A: jge 0060EA17h
  loc_0060EA0C: push 0000003Ch
  loc_0060EA0E: push 00439BF8h
  loc_0060EA13: push esi
  loc_0060EA14: push eax
  loc_0060EA15: call edi
  loc_0060EA17: lea ebx, var_30
  loc_0060EA1A: mov eax, var_2C
  loc_0060EA1D: push ebx
  loc_0060EA1E: mov edx, 00000008h
  loc_0060EA23: sub esp, 00000010h
  loc_0060EA26: mov esi, [eax]
  loc_0060EA28: mov ebx, esp
  loc_0060EA2A: mov ecx, 0043B870h ; "LabelNama"
  loc_0060EA2F: push eax
  loc_0060EA30: mov var_BC, eax
  loc_0060EA36: mov [ebx], edx
  loc_0060EA38: mov edx, var_7C
  loc_0060EA3B: mov [ebx+00000004h], edx
  loc_0060EA3E: mov [ebx+00000008h], ecx
  loc_0060EA41: mov ecx, var_74
  loc_0060EA44: mov [ebx+0000000Ch], ecx
  loc_0060EA47: call [esi+0000001Ch]
  loc_0060EA4A: test eax, eax
  loc_0060EA4C: fnclex
  loc_0060EA4E: jge 0060EA61h
  loc_0060EA50: mov edx, var_BC
  loc_0060EA56: push 0000001Ch
  loc_0060EA58: push 0043B7D8h
  loc_0060EA5D: push edx
  loc_0060EA5E: push eax
  loc_0060EA5F: call edi
  loc_0060EA61: mov ecx, [006680D8h]
  loc_0060EA67: mov edx, [006680DCh]
  loc_0060EA6D: sub esp, 00000010h
  loc_0060EA70: mov eax, esp
  loc_0060EA72: push 0043B85Ch ; "Caption"
  loc_0060EA77: mov [eax], ecx
  loc_0060EA79: mov ecx, [006680E0h]
  loc_0060EA7F: mov [eax+00000004h], edx
  loc_0060EA82: mov edx, [006680E4h]
  loc_0060EA88: mov [eax+00000008h], ecx
  loc_0060EA8B: mov [eax+0000000Ch], edx
  loc_0060EA8E: mov eax, var_30
  loc_0060EA91: push eax
  loc_0060EA92: call [00401094h] ; __vbaLateMemSt
  loc_0060EA98: lea ecx, var_30
  loc_0060EA9B: lea edx, var_2C
  loc_0060EA9E: push ecx
  loc_0060EA9F: lea eax, var_28
  loc_0060EAA2: push edx
  loc_0060EAA3: lea ecx, var_24
  loc_0060EAA6: push eax
  loc_0060EAA7: push ecx
  loc_0060EAA8: push 00000004h
  loc_0060EAAA: call [0040104Ch] ; __vbaFreeObjList
  loc_0060EAB0: mov eax, var_C4
  loc_0060EAB6: add esp, 00000014h
  loc_0060EAB9: lea ecx, var_24
  loc_0060EABC: mov edx, [eax]
  loc_0060EABE: push ecx
  loc_0060EABF: push eax
  loc_0060EAC0: call [edx+0000008Ch]
  loc_0060EAC6: test eax, eax
  loc_0060EAC8: fnclex
  loc_0060EACA: jge 0060EAE0h
  loc_0060EACC: mov edx, var_C4
  loc_0060EAD2: push 0000008Ch
  loc_0060EAD7: push 0043B5D0h
  loc_0060EADC: push edx
  loc_0060EADD: push eax
  loc_0060EADE: call edi
  loc_0060EAE0: lea ebx, var_28
  loc_0060EAE3: mov eax, var_24
  loc_0060EAE6: push ebx
  loc_0060EAE7: mov edx, 00000008h
  loc_0060EAEC: sub esp, 00000010h
  loc_0060EAEF: mov var_70, edx
  loc_0060EAF2: mov ebx, esp
  loc_0060EAF4: mov ecx, 0043B848h ; "Section4"
  loc_0060EAF9: mov var_68, ecx
  loc_0060EAFC: mov esi, [eax]
  loc_0060EAFE: mov [ebx], edx
  loc_0060EB00: mov edx, var_6C
  loc_0060EB03: push eax
  loc_0060EB04: mov var_AC, eax
  loc_0060EB0A: mov [ebx+00000004h], edx
  loc_0060EB0D: mov [ebx+00000008h], ecx
  loc_0060EB10: mov ecx, var_64
  loc_0060EB13: mov [ebx+0000000Ch], ecx
  loc_0060EB16: call [esi+0000001Ch]
  loc_0060EB19: test eax, eax
  loc_0060EB1B: fnclex
  loc_0060EB1D: jge 0060EB30h
  loc_0060EB1F: mov edx, var_AC
  loc_0060EB25: push 0000001Ch
  loc_0060EB27: push 00439DD8h
  loc_0060EB2C: push edx
  loc_0060EB2D: push eax
  loc_0060EB2E: call edi
  loc_0060EB30: mov eax, var_28
  loc_0060EB33: lea edx, var_2C
  loc_0060EB36: push edx
  loc_0060EB37: push eax
  loc_0060EB38: mov ecx, [eax]
  loc_0060EB3A: mov esi, eax
  loc_0060EB3C: call [ecx+0000003Ch]
  loc_0060EB3F: test eax, eax
  loc_0060EB41: fnclex
  loc_0060EB43: jge 0060EB50h
  loc_0060EB45: push 0000003Ch
  loc_0060EB47: push 00439BF8h
  loc_0060EB4C: push esi
  loc_0060EB4D: push eax
  loc_0060EB4E: call edi
  loc_0060EB50: lea ebx, var_30
  loc_0060EB53: mov eax, var_2C
  loc_0060EB56: push ebx
  loc_0060EB57: mov edx, 00000008h
  loc_0060EB5C: sub esp, 00000010h
  loc_0060EB5F: mov esi, [eax]
  loc_0060EB61: mov ebx, esp
  loc_0060EB63: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060EB68: push eax
  loc_0060EB69: mov var_BC, eax
  loc_0060EB6F: mov [ebx], edx
  loc_0060EB71: mov edx, var_7C
  loc_0060EB74: mov [ebx+00000004h], edx
  loc_0060EB77: mov [ebx+00000008h], ecx
  loc_0060EB7A: mov ecx, var_74
  loc_0060EB7D: mov [ebx+0000000Ch], ecx
  loc_0060EB80: call [esi+0000001Ch]
  loc_0060EB83: test eax, eax
  loc_0060EB85: fnclex
  loc_0060EB87: jge 0060EB9Ah
  loc_0060EB89: mov edx, var_BC
  loc_0060EB8F: push 0000001Ch
  loc_0060EB91: push 0043B7D8h
  loc_0060EB96: push edx
  loc_0060EB97: push eax
  loc_0060EB98: call edi
  loc_0060EB9A: mov ecx, [006680E8h]
  loc_0060EBA0: mov edx, [006680ECh]
  loc_0060EBA6: sub esp, 00000010h
  loc_0060EBA9: mov eax, esp
  loc_0060EBAB: push 0043B85Ch ; "Caption"
  loc_0060EBB0: mov [eax], ecx
  loc_0060EBB2: mov ecx, [006680F0h]
  loc_0060EBB8: mov [eax+00000004h], edx
  loc_0060EBBB: mov edx, [006680F4h]
  loc_0060EBC1: mov [eax+00000008h], ecx
  loc_0060EBC4: mov [eax+0000000Ch], edx
  loc_0060EBC7: mov eax, var_30
  loc_0060EBCA: push eax
  loc_0060EBCB: call [00401094h] ; __vbaLateMemSt
  loc_0060EBD1: lea ecx, var_30
  loc_0060EBD4: lea edx, var_2C
  loc_0060EBD7: push ecx
  loc_0060EBD8: lea eax, var_28
  loc_0060EBDB: push edx
  loc_0060EBDC: lea ecx, var_24
  loc_0060EBDF: push eax
  loc_0060EBE0: push ecx
  loc_0060EBE1: push 00000004h
  loc_0060EBE3: call [0040104Ch] ; __vbaFreeObjList
  loc_0060EBE9: mov eax, var_C4
  loc_0060EBEF: add esp, 00000014h
  loc_0060EBF2: lea ecx, var_24
  loc_0060EBF5: mov edx, [eax]
  loc_0060EBF7: push ecx
  loc_0060EBF8: push eax
  loc_0060EBF9: call [edx+0000008Ch]
  loc_0060EBFF: test eax, eax
  loc_0060EC01: fnclex
  loc_0060EC03: jge 0060EC19h
  loc_0060EC05: mov edx, var_C4
  loc_0060EC0B: push 0000008Ch
  loc_0060EC10: push 0043B5D0h
  loc_0060EC15: push edx
  loc_0060EC16: push eax
  loc_0060EC17: call edi
  loc_0060EC19: lea ebx, var_28
  loc_0060EC1C: mov eax, var_24
  loc_0060EC1F: push ebx
  loc_0060EC20: mov edx, 00000008h
  loc_0060EC25: sub esp, 00000010h
  loc_0060EC28: mov var_70, edx
  loc_0060EC2B: mov ebx, esp
  loc_0060EC2D: mov ecx, 0043B848h ; "Section4"
  loc_0060EC32: mov var_68, ecx
  loc_0060EC35: mov esi, [eax]
  loc_0060EC37: mov [ebx], edx
  loc_0060EC39: mov edx, var_6C
  loc_0060EC3C: push eax
  loc_0060EC3D: mov var_AC, eax
  loc_0060EC43: mov [ebx+00000004h], edx
  loc_0060EC46: mov [ebx+00000008h], ecx
  loc_0060EC49: mov ecx, var_64
  loc_0060EC4C: mov [ebx+0000000Ch], ecx
  loc_0060EC4F: call [esi+0000001Ch]
  loc_0060EC52: test eax, eax
  loc_0060EC54: fnclex
  loc_0060EC56: jge 0060EC69h
  loc_0060EC58: mov edx, var_AC
  loc_0060EC5E: push 0000001Ch
  loc_0060EC60: push 00439DD8h
  loc_0060EC65: push edx
  loc_0060EC66: push eax
  loc_0060EC67: call edi
  loc_0060EC69: mov eax, var_28
  loc_0060EC6C: lea edx, var_2C
  loc_0060EC6F: push edx
  loc_0060EC70: push eax
  loc_0060EC71: mov ecx, [eax]
  loc_0060EC73: mov esi, eax
  loc_0060EC75: call [ecx+0000003Ch]
  loc_0060EC78: test eax, eax
  loc_0060EC7A: fnclex
  loc_0060EC7C: jge 0060EC89h
  loc_0060EC7E: push 0000003Ch
  loc_0060EC80: push 00439BF8h
  loc_0060EC85: push esi
  loc_0060EC86: push eax
  loc_0060EC87: call edi
  loc_0060EC89: lea ebx, var_30
  loc_0060EC8C: mov eax, var_2C
  loc_0060EC8F: push ebx
  loc_0060EC90: mov edx, 00000008h
  loc_0060EC95: sub esp, 00000010h
  loc_0060EC98: mov esi, [eax]
  loc_0060EC9A: mov ebx, esp
  loc_0060EC9C: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060ECA1: push eax
  loc_0060ECA2: mov var_BC, eax
  loc_0060ECA8: mov [ebx], edx
  loc_0060ECAA: mov edx, var_7C
  loc_0060ECAD: mov [ebx+00000004h], edx
  loc_0060ECB0: mov [ebx+00000008h], ecx
  loc_0060ECB3: mov ecx, var_74
  loc_0060ECB6: mov [ebx+0000000Ch], ecx
  loc_0060ECB9: call [esi+0000001Ch]
  loc_0060ECBC: test eax, eax
  loc_0060ECBE: fnclex
  loc_0060ECC0: jge 0060ECD3h
  loc_0060ECC2: mov edx, var_BC
  loc_0060ECC8: push 0000001Ch
  loc_0060ECCA: push 0043B7D8h
  loc_0060ECCF: push edx
  loc_0060ECD0: push eax
  loc_0060ECD1: call edi
  loc_0060ECD3: mov ecx, [006680F8h]
  loc_0060ECD9: mov edx, [006680FCh]
  loc_0060ECDF: sub esp, 00000010h
  loc_0060ECE2: mov eax, esp
  loc_0060ECE4: push 0043B85Ch ; "Caption"
  loc_0060ECE9: mov [eax], ecx
  loc_0060ECEB: mov ecx, [00668100h]
  loc_0060ECF1: mov [eax+00000004h], edx
  loc_0060ECF4: mov edx, [00668104h]
  loc_0060ECFA: mov [eax+00000008h], ecx
  loc_0060ECFD: mov [eax+0000000Ch], edx
  loc_0060ED00: mov eax, var_30
  loc_0060ED03: push eax
  loc_0060ED04: call [00401094h] ; __vbaLateMemSt
  loc_0060ED0A: lea ecx, var_30
  loc_0060ED0D: lea edx, var_2C
  loc_0060ED10: push ecx
  loc_0060ED11: lea eax, var_28
  loc_0060ED14: push edx
  loc_0060ED15: lea ecx, var_24
  loc_0060ED18: push eax
  loc_0060ED19: push ecx
  loc_0060ED1A: push 00000004h
  loc_0060ED1C: call [0040104Ch] ; __vbaFreeObjList
  loc_0060ED22: mov eax, var_C4
  loc_0060ED28: add esp, 00000014h
  loc_0060ED2B: lea ecx, var_24
  loc_0060ED2E: mov edx, [eax]
  loc_0060ED30: push ecx
  loc_0060ED31: push eax
  loc_0060ED32: call [edx+0000008Ch]
  loc_0060ED38: test eax, eax
  loc_0060ED3A: fnclex
  loc_0060ED3C: jge 0060ED52h
  loc_0060ED3E: mov edx, var_C4
  loc_0060ED44: push 0000008Ch
  loc_0060ED49: push 0043B5D0h
  loc_0060ED4E: push edx
  loc_0060ED4F: push eax
  loc_0060ED50: call edi
  loc_0060ED52: lea ebx, var_28
  loc_0060ED55: mov eax, var_24
  loc_0060ED58: push ebx
  loc_0060ED59: mov edx, 00000008h
  loc_0060ED5E: sub esp, 00000010h
  loc_0060ED61: mov var_70, edx
  loc_0060ED64: mov ebx, esp
  loc_0060ED66: mov ecx, 0043B848h ; "Section4"
  loc_0060ED6B: mov var_68, ecx
  loc_0060ED6E: mov esi, [eax]
  loc_0060ED70: mov [ebx], edx
  loc_0060ED72: mov edx, var_6C
  loc_0060ED75: push eax
  loc_0060ED76: mov var_AC, eax
  loc_0060ED7C: mov [ebx+00000004h], edx
  loc_0060ED7F: mov [ebx+00000008h], ecx
  loc_0060ED82: mov ecx, var_64
  loc_0060ED85: mov [ebx+0000000Ch], ecx
  loc_0060ED88: call [esi+0000001Ch]
  loc_0060ED8B: test eax, eax
  loc_0060ED8D: fnclex
  loc_0060ED8F: jge 0060EDA2h
  loc_0060ED91: mov edx, var_AC
  loc_0060ED97: push 0000001Ch
  loc_0060ED99: push 00439DD8h
  loc_0060ED9E: push edx
  loc_0060ED9F: push eax
  loc_0060EDA0: call edi
  loc_0060EDA2: mov eax, var_28
  loc_0060EDA5: lea edx, var_2C
  loc_0060EDA8: push edx
  loc_0060EDA9: push eax
  loc_0060EDAA: mov ecx, [eax]
  loc_0060EDAC: mov esi, eax
  loc_0060EDAE: call [ecx+0000003Ch]
  loc_0060EDB1: test eax, eax
  loc_0060EDB3: fnclex
  loc_0060EDB5: jge 0060EDC2h
  loc_0060EDB7: push 0000003Ch
  loc_0060EDB9: push 00439BF8h
  loc_0060EDBE: push esi
  loc_0060EDBF: push eax
  loc_0060EDC0: call edi
  loc_0060EDC2: lea ebx, var_30
  loc_0060EDC5: mov eax, var_2C
  loc_0060EDC8: push ebx
  loc_0060EDC9: mov edx, 00000008h
  loc_0060EDCE: sub esp, 00000010h
  loc_0060EDD1: mov esi, [eax]
  loc_0060EDD3: mov ebx, esp
  loc_0060EDD5: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060EDDA: push eax
  loc_0060EDDB: mov var_BC, eax
  loc_0060EDE1: mov [ebx], edx
  loc_0060EDE3: mov edx, var_7C
  loc_0060EDE6: mov [ebx+00000004h], edx
  loc_0060EDE9: mov [ebx+00000008h], ecx
  loc_0060EDEC: mov ecx, var_74
  loc_0060EDEF: mov [ebx+0000000Ch], ecx
  loc_0060EDF2: call [esi+0000001Ch]
  loc_0060EDF5: test eax, eax
  loc_0060EDF7: fnclex
  loc_0060EDF9: jge 0060EE0Ch
  loc_0060EDFB: mov edx, var_BC
  loc_0060EE01: push 0000001Ch
  loc_0060EE03: push 0043B7D8h
  loc_0060EE08: push edx
  loc_0060EE09: push eax
  loc_0060EE0A: call edi
  loc_0060EE0C: mov ecx, [00668108h]
  loc_0060EE12: mov edx, [0066810Ch]
  loc_0060EE18: sub esp, 00000010h
  loc_0060EE1B: mov eax, esp
  loc_0060EE1D: push 0043B85Ch ; "Caption"
  loc_0060EE22: mov [eax], ecx
  loc_0060EE24: mov ecx, [00668110h]
  loc_0060EE2A: mov [eax+00000004h], edx
  loc_0060EE2D: mov edx, [00668114h]
  loc_0060EE33: mov [eax+00000008h], ecx
  loc_0060EE36: mov [eax+0000000Ch], edx
  loc_0060EE39: mov eax, var_30
  loc_0060EE3C: push eax
  loc_0060EE3D: call [00401094h] ; __vbaLateMemSt
  loc_0060EE43: lea ecx, var_30
  loc_0060EE46: lea edx, var_2C
  loc_0060EE49: push ecx
  loc_0060EE4A: lea eax, var_28
  loc_0060EE4D: push edx
  loc_0060EE4E: lea ecx, var_24
  loc_0060EE51: push eax
  loc_0060EE52: push ecx
  loc_0060EE53: push 00000004h
  loc_0060EE55: call [0040104Ch] ; __vbaFreeObjList
  loc_0060EE5B: mov eax, var_C4
  loc_0060EE61: add esp, 00000014h
  loc_0060EE64: mov edx, [eax]
  loc_0060EE66: push 0000000Ah
  loc_0060EE68: push eax
  loc_0060EE69: call [edx+00000048h]
  loc_0060EE6C: test eax, eax
  loc_0060EE6E: fnclex
  loc_0060EE70: jge 0060EE83h
  loc_0060EE72: mov ecx, var_C4
  loc_0060EE78: push 00000048h
  loc_0060EE7A: push 0043B5D0h
  loc_0060EE7F: push ecx
  loc_0060EE80: push eax
  loc_0060EE81: call edi
  loc_0060EE83: mov eax, var_C4
  loc_0060EE89: push 0000000Ah
  loc_0060EE8B: push eax
  loc_0060EE8C: mov edx, [eax]
  loc_0060EE8E: call [edx+00000058h]
  loc_0060EE91: test eax, eax
  loc_0060EE93: fnclex
  loc_0060EE95: jge 0060EEA8h
  loc_0060EE97: mov ecx, var_C4
  loc_0060EE9D: push 00000058h
  loc_0060EE9F: push 0043B5D0h
  loc_0060EEA4: push ecx
  loc_0060EEA5: push eax
  loc_0060EEA6: call edi
  loc_0060EEA8: sub esp, 00000010h
  loc_0060EEAB: mov eax, 00000002h
  loc_0060EEB0: mov edx, esp
  loc_0060EEB2: mov ecx, eax
  loc_0060EEB4: mov var_70, ecx
  loc_0060EEB7: mov var_68, eax
  loc_0060EEBA: mov [edx], ecx
  loc_0060EEBC: mov ecx, var_6C
  loc_0060EEBF: push 8001000Eh
  loc_0060EEC4: mov [edx+00000004h], ecx
  loc_0060EEC7: mov ecx, var_C4
  loc_0060EECD: push ecx
  loc_0060EECE: mov [edx+00000008h], eax
  loc_0060EED1: mov eax, var_64
  loc_0060EED4: mov [edx+0000000Ch], eax
  loc_0060EED7: call [00401280h] ; __vbaLateIdSt
  loc_0060EEDD: mov edx, var_C4
  loc_0060EEE3: push 00000000h
  loc_0060EEE5: push 80011003h
  loc_0060EEEA: push edx
  loc_0060EEEB: call [0040102Ch] ; __vbaLateIdCall
  loc_0060EEF1: add esp, 0000000Ch
  loc_0060EEF4: lea eax, var_C4
  loc_0060EEFA: push 00000000h
  loc_0060EEFC: push eax
  loc_0060EEFD: call [004010C8h] ; __vbaObjSetAddref
  loc_0060EF03: push 0060EF6Dh
  loc_0060EF08: jmp 0060EF43h
  loc_0060EF0A: lea ecx, var_20
  loc_0060EF0D: call [0040129Ch] ; __vbaFreeStr
  loc_0060EF13: lea ecx, var_30
  loc_0060EF16: lea edx, var_2C
  loc_0060EF19: push ecx
  loc_0060EF1A: lea eax, var_28
  loc_0060EF1D: push edx
  loc_0060EF1E: lea ecx, var_24
  loc_0060EF21: push eax
  loc_0060EF22: push ecx
  loc_0060EF23: push 00000004h
  loc_0060EF25: call [0040104Ch] ; __vbaFreeObjList
  loc_0060EF2B: lea edx, var_60
  loc_0060EF2E: lea eax, var_50
  loc_0060EF31: push edx
  loc_0060EF32: lea ecx, var_40
  loc_0060EF35: push eax
  loc_0060EF36: push ecx
  loc_0060EF37: push 00000003h
  loc_0060EF39: call [0040103Ch] ; __vbaFreeVarList
  loc_0060EF3F: add esp, 00000024h
  loc_0060EF42: ret
  loc_0060EF43: lea edx, var_C8
  loc_0060EF49: lea eax, var_C4
  loc_0060EF4F: push edx
  loc_0060EF50: push eax
  loc_0060EF51: push 00000002h
  loc_0060EF53: call [0040104Ch] ; __vbaFreeObjList
  loc_0060EF59: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060EF5F: add esp, 0000000Ch
  loc_0060EF62: lea ecx, var_18
  loc_0060EF65: call __vbaFreeStr
  loc_0060EF67: lea ecx, var_1C
  loc_0060EF6A: call __vbaFreeStr
  loc_0060EF6C: ret
  loc_0060EF6D: mov ecx, var_10
  loc_0060EF70: pop edi
  loc_0060EF71: pop esi
  loc_0060EF72: mov fs:[00000000h], ecx
  loc_0060EF79: pop ebx
  loc_0060EF7A: mov esp, ebp
  loc_0060EF7C: pop ebp
  loc_0060EF7D: ret
End Sub

Private Sub Proc_34_6_60EF90
  loc_0060EF90: push ebp
  loc_0060EF91: mov ebp, esp
  loc_0060EF93: sub esp, 00000008h
  loc_0060EF96: push 00405696h ; __vbaExceptHandler
  loc_0060EF9B: mov eax, fs:[00000000h]
  loc_0060EFA1: push eax
  loc_0060EFA2: mov fs:[00000000h], esp
  loc_0060EFA9: sub esp, 000000D0h
  loc_0060EFAF: push ebx
  loc_0060EFB0: push esi
  loc_0060EFB1: push edi
  loc_0060EFB2: mov var_8, esp
  loc_0060EFB5: mov var_4, 00404500h
  loc_0060EFBC: xor ebx, ebx
  loc_0060EFBE: mov var_18, ebx
  loc_0060EFC1: mov var_1C, ebx
  loc_0060EFC4: mov var_20, ebx
  loc_0060EFC7: mov var_24, ebx
  loc_0060EFCA: mov var_28, ebx
  loc_0060EFCD: mov var_2C, ebx
  loc_0060EFD0: mov var_30, ebx
  loc_0060EFD3: mov var_40, ebx
  loc_0060EFD6: mov var_50, ebx
  loc_0060EFD9: mov var_60, ebx
  loc_0060EFDC: mov var_70, ebx
  loc_0060EFDF: mov var_80, ebx
  loc_0060EFE2: mov var_90, ebx
  loc_0060EFE8: mov var_A4, ebx
  loc_0060EFEE: mov var_C4, ebx
  loc_0060EFF4: mov var_C8, ebx
  loc_0060EFFA: call 0055B320h
  loc_0060EFFF: mov edx, 0042DFECh
  loc_0060F004: lea ecx, var_18
  loc_0060F007: call [004011FCh] ; __vbaStrCopy
  loc_0060F00D: push 0043BE20h ; " SELECT no_pembelian, no_notabeli, tgl_Masuk, nama_pemasok,"
  loc_0060F012: push 0043BE9Ch ; "jenis, tot_beli FROM Pembelian order BY no_pembelian"
  loc_0060F017: call [00401070h] ; __vbaStrCat
  loc_0060F01D: mov edx, eax
  loc_0060F01F: lea ecx, var_18
  loc_0060F022: call [0040126Ch] ; __vbaStrMove
  loc_0060F028: push 0042DF30h
  loc_0060F02D: call [00401168h] ; __vbaNew
  loc_0060F033: push eax
  loc_0060F034: push 00668090h
  loc_0060F039: call [004010B8h] ; __vbaObjSet
  loc_0060F03F: cmp [0066803Ch], ebx
  loc_0060F045: jnz 0060F057h
  loc_0060F047: push 0066803Ch
  loc_0060F04C: push 0042DEFCh
  loc_0060F051: call [004011E8h] ; __vbaNew2
  loc_0060F057: mov esi, [0066803Ch]
  loc_0060F05D: lea ecx, var_40
  loc_0060F060: mov var_38, 80020004h
  loc_0060F067: mov var_40, 0000000Ah
  loc_0060F06E: call [0040123Ch] ; __vbaFreeVarg
  loc_0060F074: mov eax, [esi]
  loc_0060F076: lea ecx, var_24
  loc_0060F079: push ecx
  loc_0060F07A: mov ecx, var_18
  loc_0060F07D: lea edx, var_40
  loc_0060F080: push 00000001h
  loc_0060F082: push edx
  loc_0060F083: push ecx
  loc_0060F084: push esi
  loc_0060F085: call [eax+00000040h]
  loc_0060F088: cmp eax, ebx
  loc_0060F08A: fnclex
  loc_0060F08C: jge 0060F0A1h
  loc_0060F08E: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060F094: push 00000040h
  loc_0060F096: push 0042E1B0h
  loc_0060F09B: push esi
  loc_0060F09C: push eax
  loc_0060F09D: call edi
  loc_0060F09F: jmp 0060F0A7h
  loc_0060F0A1: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060F0A7: mov edx, var_24
  loc_0060F0AA: mov esi, [00401268h] ; __vbaCastObj
  loc_0060F0B0: push 0042DF20h
  loc_0060F0B5: push edx
  loc_0060F0B6: call __vbaCastObj
  loc_0060F0B8: push eax
  loc_0060F0B9: push 00668090h
  loc_0060F0BE: call [004010B8h] ; __vbaObjSet
  loc_0060F0C4: lea ecx, var_24
  loc_0060F0C7: call [004012A0h] ; __vbaFreeObj
  loc_0060F0CD: lea ecx, var_40
  loc_0060F0D0: call [00401020h] ; __vbaFreeVar
  loc_0060F0D6: cmp [0066856Ch], ebx
  loc_0060F0DC: jnz 0060F0EEh
  loc_0060F0DE: push 0066856Ch
  loc_0060F0E3: push 00407230h
  loc_0060F0E8: call [004011E8h] ; __vbaNew2
  loc_0060F0EE: mov eax, [0066856Ch]
  loc_0060F0F3: lea ecx, var_C4
  loc_0060F0F9: push eax
  loc_0060F0FA: push ecx
  loc_0060F0FB: call [004010C8h] ; __vbaObjSetAddref
  loc_0060F101: mov edx, var_C4
  loc_0060F107: push 0043B7B8h
  loc_0060F10C: push ebx
  loc_0060F10D: mov edx, [edx]
  loc_0060F10F: mov var_E0, edx
  loc_0060F115: call __vbaCastObj
  loc_0060F117: push eax
  loc_0060F118: lea eax, var_24
  loc_0060F11B: push eax
  loc_0060F11C: call [004010B8h] ; __vbaObjSet
  loc_0060F122: mov ecx, var_C4
  loc_0060F128: mov edx, var_E0
  loc_0060F12E: push eax
  loc_0060F12F: push ecx
  loc_0060F130: call [edx+00000028h]
  loc_0060F133: cmp eax, ebx
  loc_0060F135: fnclex
  loc_0060F137: jge 0060F14Ah
  loc_0060F139: mov ecx, var_C4
  loc_0060F13F: push 00000028h
  loc_0060F141: push 0043B5D0h
  loc_0060F146: push ecx
  loc_0060F147: push eax
  loc_0060F148: call edi
  loc_0060F14A: lea ecx, var_24
  loc_0060F14D: call [004012A0h] ; __vbaFreeObj
  loc_0060F153: mov eax, var_C4
  loc_0060F159: push 0042DFECh
  loc_0060F15E: push eax
  loc_0060F15F: mov edx, [eax]
  loc_0060F161: call [edx+00000030h]
  loc_0060F164: cmp eax, ebx
  loc_0060F166: fnclex
  loc_0060F168: jge 0060F17Bh
  loc_0060F16A: mov ecx, var_C4
  loc_0060F170: push 00000030h
  loc_0060F172: push 0043B5D0h
  loc_0060F177: push ecx
  loc_0060F178: push eax
  loc_0060F179: call edi
  loc_0060F17B: mov eax, [00668090h]
  loc_0060F180: lea ecx, var_24
  loc_0060F183: push ecx
  loc_0060F184: push eax
  loc_0060F185: mov edx, [eax]
  loc_0060F187: call [edx+00000114h]
  loc_0060F18D: cmp eax, ebx
  loc_0060F18F: fnclex
  loc_0060F191: jge 0060F1A7h
  loc_0060F193: mov edx, [00668090h]
  loc_0060F199: push 00000114h
  loc_0060F19E: push 0043B7C8h
  loc_0060F1A3: push edx
  loc_0060F1A4: push eax
  loc_0060F1A5: call edi
  loc_0060F1A7: mov eax, var_C4
  loc_0060F1AD: mov ecx, var_24
  loc_0060F1B0: push 0043B7B8h
  loc_0060F1B5: push ecx
  loc_0060F1B6: mov esi, [eax]
  loc_0060F1B8: call [00401268h] ; __vbaCastObj
  loc_0060F1BE: lea edx, var_28
  loc_0060F1C1: push eax
  loc_0060F1C2: push edx
  loc_0060F1C3: call [004010B8h] ; __vbaObjSet
  loc_0060F1C9: push eax
  loc_0060F1CA: mov eax, var_C4
  loc_0060F1D0: push eax
  loc_0060F1D1: call [esi+00000028h]
  loc_0060F1D4: cmp eax, ebx
  loc_0060F1D6: fnclex
  loc_0060F1D8: jge 0060F1EBh
  loc_0060F1DA: mov ecx, var_C4
  loc_0060F1E0: push 00000028h
  loc_0060F1E2: push 0043B5D0h
  loc_0060F1E7: push ecx
  loc_0060F1E8: push eax
  loc_0060F1E9: call edi
  loc_0060F1EB: lea edx, var_28
  loc_0060F1EE: lea eax, var_24
  loc_0060F1F1: push edx
  loc_0060F1F2: push eax
  loc_0060F1F3: push 00000002h
  loc_0060F1F5: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F1FB: mov eax, var_C4
  loc_0060F201: add esp, 0000000Ch
  loc_0060F204: lea edx, var_24
  loc_0060F207: mov ecx, [eax]
  loc_0060F209: push edx
  loc_0060F20A: push eax
  loc_0060F20B: call [ecx+0000008Ch]
  loc_0060F211: cmp eax, ebx
  loc_0060F213: fnclex
  loc_0060F215: jge 0060F22Bh
  loc_0060F217: mov ecx, var_C4
  loc_0060F21D: push 0000008Ch
  loc_0060F222: push 0043B5D0h
  loc_0060F227: push ecx
  loc_0060F228: push eax
  loc_0060F229: call edi
  loc_0060F22B: lea esi, var_28
  loc_0060F22E: mov eax, var_24
  loc_0060F231: push esi
  loc_0060F232: mov ecx, 00000008h
  loc_0060F237: sub esp, 00000010h
  loc_0060F23A: mov var_70, ecx
  loc_0060F23D: mov esi, esp
  loc_0060F23F: mov var_68, 00439FB0h ; "Section1"
  loc_0060F246: mov edx, [eax]
  loc_0060F248: push eax
  loc_0060F249: mov [esi], ecx
  loc_0060F24B: mov ecx, var_6C
  loc_0060F24E: mov var_AC, eax
  loc_0060F254: mov [esi+00000004h], ecx
  loc_0060F257: mov ecx, var_68
  loc_0060F25A: mov [esi+00000008h], ecx
  loc_0060F25D: mov ecx, var_64
  loc_0060F260: mov [esi+0000000Ch], ecx
  loc_0060F263: call [edx+0000001Ch]
  loc_0060F266: cmp eax, ebx
  loc_0060F268: fnclex
  loc_0060F26A: jge 0060F27Dh
  loc_0060F26C: mov edx, var_AC
  loc_0060F272: push 0000001Ch
  loc_0060F274: push 00439DD8h
  loc_0060F279: push edx
  loc_0060F27A: push eax
  loc_0060F27B: call edi
  loc_0060F27D: mov eax, var_28
  loc_0060F280: lea edx, var_2C
  loc_0060F283: push edx
  loc_0060F284: push eax
  loc_0060F285: mov ecx, [eax]
  loc_0060F287: mov esi, eax
  loc_0060F289: call [ecx+0000003Ch]
  loc_0060F28C: cmp eax, ebx
  loc_0060F28E: fnclex
  loc_0060F290: jge 0060F29Dh
  loc_0060F292: push 0000003Ch
  loc_0060F294: push 00439BF8h
  loc_0060F299: push esi
  loc_0060F29A: push eax
  loc_0060F29B: call edi
  loc_0060F29D: mov eax, var_2C
  loc_0060F2A0: mov var_2C, ebx
  loc_0060F2A3: push eax
  loc_0060F2A4: lea eax, var_C8
  loc_0060F2AA: push eax
  loc_0060F2AB: call [004010B8h] ; __vbaObjSet
  loc_0060F2B1: lea ecx, var_28
  loc_0060F2B4: lea edx, var_24
  loc_0060F2B7: push ecx
  loc_0060F2B8: push edx
  loc_0060F2B9: push 00000002h
  loc_0060F2BB: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F2C1: mov eax, var_C8
  loc_0060F2C7: add esp, 0000000Ch
  loc_0060F2CA: lea edx, var_A4
  loc_0060F2D0: mov ecx, [eax]
  loc_0060F2D2: push edx
  loc_0060F2D3: push eax
  loc_0060F2D4: call [ecx+00000020h]
  loc_0060F2D7: cmp eax, ebx
  loc_0060F2D9: fnclex
  loc_0060F2DB: jge 0060F2EEh
  loc_0060F2DD: mov ecx, var_C8
  loc_0060F2E3: push 00000020h
  loc_0060F2E5: push 0043B7D8h
  loc_0060F2EA: push ecx
  loc_0060F2EB: push eax
  loc_0060F2EC: call edi
  loc_0060F2EE: mov ecx, var_A4
  loc_0060F2F4: call [00401138h] ; __vbaI2I4
  loc_0060F2FA: mov var_D0, eax
  loc_0060F300: mov ebx, 00000001h
  loc_0060F305: cmp bx, var_D0
  loc_0060F30C: mov var_14, ebx
  loc_0060F30F: jg 0060F580h
  loc_0060F315: lea esi, var_24
  loc_0060F318: mov ecx, var_C8
  loc_0060F31E: push esi
  loc_0060F31F: mov eax, 00000002h
  loc_0060F324: sub esp, 00000010h
  loc_0060F327: mov var_70, eax
  loc_0060F32A: mov esi, esp
  loc_0060F32C: mov var_68, bx
  loc_0060F330: mov edx, [ecx]
  loc_0060F332: push ecx
  loc_0060F333: mov [esi], eax
  loc_0060F335: mov eax, var_6C
  loc_0060F338: mov [esi+00000004h], eax
  loc_0060F33B: mov eax, var_68
  loc_0060F33E: mov [esi+00000008h], eax
  loc_0060F341: mov eax, var_64
  loc_0060F344: mov [esi+0000000Ch], eax
  loc_0060F347: call [edx+0000001Ch]
  loc_0060F34A: test eax, eax
  loc_0060F34C: fnclex
  loc_0060F34E: jge 0060F361h
  loc_0060F350: mov ecx, var_C8
  loc_0060F356: push 0000001Ch
  loc_0060F358: push 0043B7D8h
  loc_0060F35D: push ecx
  loc_0060F35E: push eax
  loc_0060F35F: call edi
  loc_0060F361: mov edx, var_24
  loc_0060F364: push 0043B7E8h
  loc_0060F369: push edx
  loc_0060F36A: call [004011CCh] ; __vbaCheckType
  loc_0060F370: lea ecx, var_24
  loc_0060F373: mov si, ax
  loc_0060F376: call [004012A0h] ; __vbaFreeObj
  loc_0060F37C: test si, si
  loc_0060F37F: Unknown_795()
  loc_0060F385: mov var_68, bx
  loc_0060F389: lea ebx, var_24
  loc_0060F38C: push ebx
  loc_0060F38D: mov ecx, var_C8
  loc_0060F393: sub esp, 00000010h
  loc_0060F396: mov eax, 00000002h
  loc_0060F39B: mov ebx, esp
  loc_0060F39D: mov var_70, eax
  loc_0060F3A0: mov edx, [ecx]
  loc_0060F3A2: mov esi, 0042DFECh
  loc_0060F3A7: mov [ebx], eax
  loc_0060F3A9: mov eax, var_6C
  loc_0060F3AC: push ecx
  loc_0060F3AD: mov var_78, esi
  loc_0060F3B0: mov [ebx+00000004h], eax
  loc_0060F3B3: mov eax, var_68
  loc_0060F3B6: mov edi, 00000008h
  loc_0060F3BB: mov [ebx+00000008h], eax
  loc_0060F3BE: mov eax, var_64
  loc_0060F3C1: mov [ebx+0000000Ch], eax
  loc_0060F3C4: call [edx+0000001Ch]
  loc_0060F3C7: test eax, eax
  loc_0060F3C9: fnclex
  loc_0060F3CB: jge 0060F3E2h
  loc_0060F3CD: mov ecx, var_C8
  loc_0060F3D3: push 0000001Ch
  loc_0060F3D5: push 0043B7D8h
  loc_0060F3DA: push ecx
  loc_0060F3DB: push eax
  loc_0060F3DC: call [00401080h] ; __vbaHresultCheckObj
  loc_0060F3E2: mov eax, var_7C
  loc_0060F3E5: sub esp, 00000010h
  loc_0060F3E8: mov edx, esp
  loc_0060F3EA: mov ecx, var_74
  loc_0060F3ED: push 0043B7F8h ; "DataMember"
  loc_0060F3F2: mov [edx], edi
  loc_0060F3F4: mov [edx+00000004h], eax
  loc_0060F3F7: mov [edx+00000008h], esi
  loc_0060F3FA: mov [edx+0000000Ch], ecx
  loc_0060F3FD: mov edx, var_24
  loc_0060F400: push edx
  loc_0060F401: call [00401094h] ; __vbaLateMemSt
  loc_0060F407: lea ecx, var_24
  loc_0060F40A: call [004012A0h] ; __vbaFreeObj
  loc_0060F410: mov eax, [00668090h]
  loc_0060F415: lea edx, var_24
  loc_0060F418: push edx
  loc_0060F419: push eax
  loc_0060F41A: mov ecx, [eax]
  loc_0060F41C: call [ecx+00000054h]
  loc_0060F41F: test eax, eax
  loc_0060F421: fnclex
  loc_0060F423: jge 0060F43Ah
  loc_0060F425: mov ecx, [00668090h]
  loc_0060F42B: push 00000054h
  loc_0060F42D: push 0042DF44h
  loc_0060F432: push ecx
  loc_0060F433: push eax
  loc_0060F434: call [00401080h] ; __vbaHresultCheckObj
  loc_0060F43A: mov ebx, var_14
  loc_0060F43D: lea edi, var_28
  loc_0060F440: mov dx, bx
  loc_0060F443: push edi
  loc_0060F444: sub dx, 0001h
  loc_0060F448: mov eax, var_24
  loc_0060F44B: jo 0060FD0Eh
  loc_0060F451: sub esp, 00000010h
  loc_0060F454: mov ecx, 00000002h
  loc_0060F459: mov edi, esp
  loc_0060F45B: mov var_70, ecx
  loc_0060F45E: mov var_68, dx
  loc_0060F462: mov edx, [eax]
  loc_0060F464: mov [edi], ecx
  loc_0060F466: mov ecx, var_6C
  loc_0060F469: push eax
  loc_0060F46A: mov esi, eax
  loc_0060F46C: mov [edi+00000004h], ecx
  loc_0060F46F: mov ecx, var_68
  loc_0060F472: mov [edi+00000008h], ecx
  loc_0060F475: mov ecx, var_64
  loc_0060F478: mov [edi+0000000Ch], ecx
  loc_0060F47B: call [edx+00000028h]
  loc_0060F47E: test eax, eax
  loc_0060F480: fnclex
  loc_0060F482: jge 0060F497h
  loc_0060F484: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060F48A: push 00000028h
  loc_0060F48C: push 0042DFACh
  loc_0060F491: push esi
  loc_0060F492: push eax
  loc_0060F493: call edi
  loc_0060F495: jmp 0060F49Dh
  loc_0060F497: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060F49D: mov eax, var_28
  loc_0060F4A0: lea ecx, var_20
  loc_0060F4A3: push ecx
  loc_0060F4A4: push eax
  loc_0060F4A5: mov edx, [eax]
  loc_0060F4A7: mov esi, eax
  loc_0060F4A9: call [edx+0000002Ch]
  loc_0060F4AC: test eax, eax
  loc_0060F4AE: fnclex
  loc_0060F4B0: jge 0060F4BDh
  loc_0060F4B2: push 0000002Ch
  loc_0060F4B4: push 0042DFBCh
  loc_0060F4B9: push esi
  loc_0060F4BA: push eax
  loc_0060F4BB: call edi
  loc_0060F4BD: mov eax, var_20
  loc_0060F4C0: lea esi, var_2C
  loc_0060F4C3: push esi
  loc_0060F4C4: mov ecx, var_C8
  loc_0060F4CA: sub esp, 00000010h
  loc_0060F4CD: mov var_38, eax
  loc_0060F4D0: mov esi, esp
  loc_0060F4D2: mov eax, 00000002h
  loc_0060F4D7: mov var_78, bx
  loc_0060F4DB: mov var_20, 00000000h
  loc_0060F4E2: mov [esi], eax
  loc_0060F4E4: mov eax, var_7C
  loc_0060F4E7: mov var_40, 00000008h
  loc_0060F4EE: mov edx, [ecx]
  loc_0060F4F0: mov [esi+00000004h], eax
  loc_0060F4F3: mov eax, var_78
  loc_0060F4F6: push ecx
  loc_0060F4F7: mov [esi+00000008h], eax
  loc_0060F4FA: mov eax, var_74
  loc_0060F4FD: mov [esi+0000000Ch], eax
  loc_0060F500: call [edx+0000001Ch]
  loc_0060F503: test eax, eax
  loc_0060F505: fnclex
  loc_0060F507: jge 0060F51Ah
  loc_0060F509: mov ecx, var_C8
  loc_0060F50F: push 0000001Ch
  loc_0060F511: push 0043B7D8h
  loc_0060F516: push ecx
  loc_0060F517: push eax
  loc_0060F518: call edi
  loc_0060F51A: mov eax, var_40
  loc_0060F51D: mov ecx, var_3C
  loc_0060F520: sub esp, 00000010h
  loc_0060F523: mov edx, esp
  loc_0060F525: push 0043B810h ; "DataField"
  loc_0060F52A: mov [edx], eax
  loc_0060F52C: mov eax, var_38
  loc_0060F52F: mov [edx+00000004h], ecx
  loc_0060F532: mov ecx, var_34
  loc_0060F535: mov [edx+00000008h], eax
  loc_0060F538: mov [edx+0000000Ch], ecx
  loc_0060F53B: mov edx, var_2C
  loc_0060F53E: push edx
  loc_0060F53F: call [00401094h] ; __vbaLateMemSt
  loc_0060F545: lea eax, var_2C
  loc_0060F548: lea ecx, var_28
  loc_0060F54B: push eax
  loc_0060F54C: lea edx, var_24
  loc_0060F54F: push ecx
  loc_0060F550: push edx
  loc_0060F551: push 00000003h
  loc_0060F553: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F559: add esp, 00000010h
  loc_0060F55C: lea ecx, var_40
  loc_0060F55F: call [00401020h] ; __vbaFreeVar
  loc_0060F565: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060F56B: mov eax, 00000001h
  loc_0060F570: add ax, bx
  loc_0060F573: jo 0060FD0Eh
  loc_0060F579: mov ebx, eax
  loc_0060F57B: jmp 0060F305h
  loc_0060F580: lea eax, var_C8
  loc_0060F586: push 00000000h
  loc_0060F588: push eax
  loc_0060F589: call [004010C8h] ; __vbaObjSetAddref
  loc_0060F58F: lea ecx, var_40
  loc_0060F592: push ecx
  loc_0060F593: call [00401220h] ; rtcGetDateVar
  loc_0060F599: lea edx, var_70
  loc_0060F59C: lea ecx, var_50
  loc_0060F59F: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_0060F5A6: mov var_70, 00000008h
  loc_0060F5AD: call [00401238h] ; __vbaVarDup
  loc_0060F5B3: push 00000001h
  loc_0060F5B5: lea edx, var_50
  loc_0060F5B8: push 00000001h
  loc_0060F5BA: lea eax, var_40
  loc_0060F5BD: push edx
  loc_0060F5BE: lea ecx, var_60
  loc_0060F5C1: push eax
  loc_0060F5C2: push ecx
  loc_0060F5C3: call [00401078h] ; rtcVarFromFormatVar
  loc_0060F5C9: mov eax, var_C4
  loc_0060F5CF: lea ecx, var_24
  loc_0060F5D2: push ecx
  loc_0060F5D3: push eax
  loc_0060F5D4: mov edx, [eax]
  loc_0060F5D6: call [edx+0000008Ch]
  loc_0060F5DC: test eax, eax
  loc_0060F5DE: fnclex
  loc_0060F5E0: jge 0060F5F6h
  loc_0060F5E2: mov edx, var_C4
  loc_0060F5E8: push 0000008Ch
  loc_0060F5ED: push 0043B5D0h
  loc_0060F5F2: push edx
  loc_0060F5F3: push eax
  loc_0060F5F4: call edi
  loc_0060F5F6: lea ebx, var_28
  loc_0060F5F9: mov eax, var_24
  loc_0060F5FC: push ebx
  loc_0060F5FD: mov edx, 00000008h
  loc_0060F602: sub esp, 00000010h
  loc_0060F605: mov esi, [eax]
  loc_0060F607: mov ebx, esp
  loc_0060F609: mov ecx, 0043B848h ; "Section4"
  loc_0060F60E: push eax
  loc_0060F60F: mov var_AC, eax
  loc_0060F615: mov [ebx], edx
  loc_0060F617: mov edx, var_7C
  loc_0060F61A: mov [ebx+00000004h], edx
  loc_0060F61D: mov [ebx+00000008h], ecx
  loc_0060F620: mov ecx, var_74
  loc_0060F623: mov [ebx+0000000Ch], ecx
  loc_0060F626: call [esi+0000001Ch]
  loc_0060F629: test eax, eax
  loc_0060F62B: fnclex
  loc_0060F62D: jge 0060F640h
  loc_0060F62F: mov edx, var_AC
  loc_0060F635: push 0000001Ch
  loc_0060F637: push 00439DD8h
  loc_0060F63C: push edx
  loc_0060F63D: push eax
  loc_0060F63E: call edi
  loc_0060F640: mov eax, var_28
  loc_0060F643: lea edx, var_2C
  loc_0060F646: push edx
  loc_0060F647: push eax
  loc_0060F648: mov ecx, [eax]
  loc_0060F64A: mov esi, eax
  loc_0060F64C: call [ecx+0000003Ch]
  loc_0060F64F: test eax, eax
  loc_0060F651: fnclex
  loc_0060F653: jge 0060F660h
  loc_0060F655: push 0000003Ch
  loc_0060F657: push 00439BF8h
  loc_0060F65C: push esi
  loc_0060F65D: push eax
  loc_0060F65E: call edi
  loc_0060F660: lea ebx, var_30
  loc_0060F663: mov eax, var_2C
  loc_0060F666: push ebx
  loc_0060F667: mov edx, 00000008h
  loc_0060F66C: sub esp, 00000010h
  loc_0060F66F: mov esi, [eax]
  loc_0060F671: mov ebx, esp
  loc_0060F673: mov ecx, 0043B828h ; "labelTanggal"
  loc_0060F678: push eax
  loc_0060F679: mov var_BC, eax
  loc_0060F67F: mov [ebx], edx
  loc_0060F681: mov edx, var_8C
  loc_0060F687: mov [ebx+00000004h], edx
  loc_0060F68A: mov [ebx+00000008h], ecx
  loc_0060F68D: mov ecx, var_84
  loc_0060F693: mov [ebx+0000000Ch], ecx
  loc_0060F696: call [esi+0000001Ch]
  loc_0060F699: test eax, eax
  loc_0060F69B: fnclex
  loc_0060F69D: jge 0060F6B0h
  loc_0060F69F: mov edx, var_BC
  loc_0060F6A5: push 0000001Ch
  loc_0060F6A7: push 0043B7D8h
  loc_0060F6AC: push edx
  loc_0060F6AD: push eax
  loc_0060F6AE: call edi
  loc_0060F6B0: mov ecx, var_60
  loc_0060F6B3: mov edx, var_5C
  loc_0060F6B6: sub esp, 00000010h
  loc_0060F6B9: mov eax, esp
  loc_0060F6BB: push 0043B85Ch ; "Caption"
  loc_0060F6C0: mov [eax], ecx
  loc_0060F6C2: mov ecx, var_58
  loc_0060F6C5: mov [eax+00000004h], edx
  loc_0060F6C8: mov edx, var_54
  loc_0060F6CB: mov [eax+00000008h], ecx
  loc_0060F6CE: mov [eax+0000000Ch], edx
  loc_0060F6D1: mov eax, var_30
  loc_0060F6D4: push eax
  loc_0060F6D5: call [00401094h] ; __vbaLateMemSt
  loc_0060F6DB: lea ecx, var_30
  loc_0060F6DE: lea edx, var_2C
  loc_0060F6E1: push ecx
  loc_0060F6E2: lea eax, var_28
  loc_0060F6E5: push edx
  loc_0060F6E6: lea ecx, var_24
  loc_0060F6E9: push eax
  loc_0060F6EA: push ecx
  loc_0060F6EB: push 00000004h
  loc_0060F6ED: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F6F3: lea edx, var_60
  loc_0060F6F6: lea eax, var_50
  loc_0060F6F9: push edx
  loc_0060F6FA: lea ecx, var_40
  loc_0060F6FD: push eax
  loc_0060F6FE: push ecx
  loc_0060F6FF: push 00000003h
  loc_0060F701: call [0040103Ch] ; __vbaFreeVarList
  loc_0060F707: mov eax, var_C4
  loc_0060F70D: add esp, 00000024h
  loc_0060F710: lea ecx, var_24
  loc_0060F713: mov edx, [eax]
  loc_0060F715: push ecx
  loc_0060F716: push eax
  loc_0060F717: call [edx+0000008Ch]
  loc_0060F71D: test eax, eax
  loc_0060F71F: fnclex
  loc_0060F721: jge 0060F737h
  loc_0060F723: mov edx, var_C4
  loc_0060F729: push 0000008Ch
  loc_0060F72E: push 0043B5D0h
  loc_0060F733: push edx
  loc_0060F734: push eax
  loc_0060F735: call edi
  loc_0060F737: lea ebx, var_28
  loc_0060F73A: mov eax, var_24
  loc_0060F73D: push ebx
  loc_0060F73E: mov edx, 00000008h
  loc_0060F743: sub esp, 00000010h
  loc_0060F746: mov var_70, edx
  loc_0060F749: mov ebx, esp
  loc_0060F74B: mov ecx, 0043B848h ; "Section4"
  loc_0060F750: mov var_68, ecx
  loc_0060F753: mov esi, [eax]
  loc_0060F755: mov [ebx], edx
  loc_0060F757: mov edx, var_6C
  loc_0060F75A: push eax
  loc_0060F75B: mov var_AC, eax
  loc_0060F761: mov [ebx+00000004h], edx
  loc_0060F764: mov [ebx+00000008h], ecx
  loc_0060F767: mov ecx, var_64
  loc_0060F76A: mov [ebx+0000000Ch], ecx
  loc_0060F76D: call [esi+0000001Ch]
  loc_0060F770: test eax, eax
  loc_0060F772: fnclex
  loc_0060F774: jge 0060F787h
  loc_0060F776: mov edx, var_AC
  loc_0060F77C: push 0000001Ch
  loc_0060F77E: push 00439DD8h
  loc_0060F783: push edx
  loc_0060F784: push eax
  loc_0060F785: call edi
  loc_0060F787: mov eax, var_28
  loc_0060F78A: lea edx, var_2C
  loc_0060F78D: push edx
  loc_0060F78E: push eax
  loc_0060F78F: mov ecx, [eax]
  loc_0060F791: mov esi, eax
  loc_0060F793: call [ecx+0000003Ch]
  loc_0060F796: test eax, eax
  loc_0060F798: fnclex
  loc_0060F79A: jge 0060F7A7h
  loc_0060F79C: push 0000003Ch
  loc_0060F79E: push 00439BF8h
  loc_0060F7A3: push esi
  loc_0060F7A4: push eax
  loc_0060F7A5: call edi
  loc_0060F7A7: lea ebx, var_30
  loc_0060F7AA: mov eax, var_2C
  loc_0060F7AD: push ebx
  loc_0060F7AE: mov edx, 00000008h
  loc_0060F7B3: sub esp, 00000010h
  loc_0060F7B6: mov esi, [eax]
  loc_0060F7B8: mov ebx, esp
  loc_0060F7BA: mov ecx, 0043B870h ; "LabelNama"
  loc_0060F7BF: push eax
  loc_0060F7C0: mov var_BC, eax
  loc_0060F7C6: mov [ebx], edx
  loc_0060F7C8: mov edx, var_7C
  loc_0060F7CB: mov [ebx+00000004h], edx
  loc_0060F7CE: mov [ebx+00000008h], ecx
  loc_0060F7D1: mov ecx, var_74
  loc_0060F7D4: mov [ebx+0000000Ch], ecx
  loc_0060F7D7: call [esi+0000001Ch]
  loc_0060F7DA: test eax, eax
  loc_0060F7DC: fnclex
  loc_0060F7DE: jge 0060F7F1h
  loc_0060F7E0: mov edx, var_BC
  loc_0060F7E6: push 0000001Ch
  loc_0060F7E8: push 0043B7D8h
  loc_0060F7ED: push edx
  loc_0060F7EE: push eax
  loc_0060F7EF: call edi
  loc_0060F7F1: mov ecx, [006680D8h]
  loc_0060F7F7: mov edx, [006680DCh]
  loc_0060F7FD: sub esp, 00000010h
  loc_0060F800: mov eax, esp
  loc_0060F802: push 0043B85Ch ; "Caption"
  loc_0060F807: mov [eax], ecx
  loc_0060F809: mov ecx, [006680E0h]
  loc_0060F80F: mov [eax+00000004h], edx
  loc_0060F812: mov edx, [006680E4h]
  loc_0060F818: mov [eax+00000008h], ecx
  loc_0060F81B: mov [eax+0000000Ch], edx
  loc_0060F81E: mov eax, var_30
  loc_0060F821: push eax
  loc_0060F822: call [00401094h] ; __vbaLateMemSt
  loc_0060F828: lea ecx, var_30
  loc_0060F82B: lea edx, var_2C
  loc_0060F82E: push ecx
  loc_0060F82F: lea eax, var_28
  loc_0060F832: push edx
  loc_0060F833: lea ecx, var_24
  loc_0060F836: push eax
  loc_0060F837: push ecx
  loc_0060F838: push 00000004h
  loc_0060F83A: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F840: mov eax, var_C4
  loc_0060F846: add esp, 00000014h
  loc_0060F849: lea ecx, var_24
  loc_0060F84C: mov edx, [eax]
  loc_0060F84E: push ecx
  loc_0060F84F: push eax
  loc_0060F850: call [edx+0000008Ch]
  loc_0060F856: test eax, eax
  loc_0060F858: fnclex
  loc_0060F85A: jge 0060F870h
  loc_0060F85C: mov edx, var_C4
  loc_0060F862: push 0000008Ch
  loc_0060F867: push 0043B5D0h
  loc_0060F86C: push edx
  loc_0060F86D: push eax
  loc_0060F86E: call edi
  loc_0060F870: lea ebx, var_28
  loc_0060F873: mov eax, var_24
  loc_0060F876: push ebx
  loc_0060F877: mov edx, 00000008h
  loc_0060F87C: sub esp, 00000010h
  loc_0060F87F: mov var_70, edx
  loc_0060F882: mov ebx, esp
  loc_0060F884: mov ecx, 0043B848h ; "Section4"
  loc_0060F889: mov var_68, ecx
  loc_0060F88C: mov esi, [eax]
  loc_0060F88E: mov [ebx], edx
  loc_0060F890: mov edx, var_6C
  loc_0060F893: push eax
  loc_0060F894: mov var_AC, eax
  loc_0060F89A: mov [ebx+00000004h], edx
  loc_0060F89D: mov [ebx+00000008h], ecx
  loc_0060F8A0: mov ecx, var_64
  loc_0060F8A3: mov [ebx+0000000Ch], ecx
  loc_0060F8A6: call [esi+0000001Ch]
  loc_0060F8A9: test eax, eax
  loc_0060F8AB: fnclex
  loc_0060F8AD: jge 0060F8C0h
  loc_0060F8AF: mov edx, var_AC
  loc_0060F8B5: push 0000001Ch
  loc_0060F8B7: push 00439DD8h
  loc_0060F8BC: push edx
  loc_0060F8BD: push eax
  loc_0060F8BE: call edi
  loc_0060F8C0: mov eax, var_28
  loc_0060F8C3: lea edx, var_2C
  loc_0060F8C6: push edx
  loc_0060F8C7: push eax
  loc_0060F8C8: mov ecx, [eax]
  loc_0060F8CA: mov esi, eax
  loc_0060F8CC: call [ecx+0000003Ch]
  loc_0060F8CF: test eax, eax
  loc_0060F8D1: fnclex
  loc_0060F8D3: jge 0060F8E0h
  loc_0060F8D5: push 0000003Ch
  loc_0060F8D7: push 00439BF8h
  loc_0060F8DC: push esi
  loc_0060F8DD: push eax
  loc_0060F8DE: call edi
  loc_0060F8E0: lea ebx, var_30
  loc_0060F8E3: mov eax, var_2C
  loc_0060F8E6: push ebx
  loc_0060F8E7: mov edx, 00000008h
  loc_0060F8EC: sub esp, 00000010h
  loc_0060F8EF: mov esi, [eax]
  loc_0060F8F1: mov ebx, esp
  loc_0060F8F3: mov ecx, 0043B888h ; "LabelAlamat"
  loc_0060F8F8: push eax
  loc_0060F8F9: mov var_BC, eax
  loc_0060F8FF: mov [ebx], edx
  loc_0060F901: mov edx, var_7C
  loc_0060F904: mov [ebx+00000004h], edx
  loc_0060F907: mov [ebx+00000008h], ecx
  loc_0060F90A: mov ecx, var_74
  loc_0060F90D: mov [ebx+0000000Ch], ecx
  loc_0060F910: call [esi+0000001Ch]
  loc_0060F913: test eax, eax
  loc_0060F915: fnclex
  loc_0060F917: jge 0060F92Ah
  loc_0060F919: mov edx, var_BC
  loc_0060F91F: push 0000001Ch
  loc_0060F921: push 0043B7D8h
  loc_0060F926: push edx
  loc_0060F927: push eax
  loc_0060F928: call edi
  loc_0060F92A: mov ecx, [006680E8h]
  loc_0060F930: mov edx, [006680ECh]
  loc_0060F936: sub esp, 00000010h
  loc_0060F939: mov eax, esp
  loc_0060F93B: push 0043B85Ch ; "Caption"
  loc_0060F940: mov [eax], ecx
  loc_0060F942: mov ecx, [006680F0h]
  loc_0060F948: mov [eax+00000004h], edx
  loc_0060F94B: mov edx, [006680F4h]
  loc_0060F951: mov [eax+00000008h], ecx
  loc_0060F954: mov [eax+0000000Ch], edx
  loc_0060F957: mov eax, var_30
  loc_0060F95A: push eax
  loc_0060F95B: call [00401094h] ; __vbaLateMemSt
  loc_0060F961: lea ecx, var_30
  loc_0060F964: lea edx, var_2C
  loc_0060F967: push ecx
  loc_0060F968: lea eax, var_28
  loc_0060F96B: push edx
  loc_0060F96C: lea ecx, var_24
  loc_0060F96F: push eax
  loc_0060F970: push ecx
  loc_0060F971: push 00000004h
  loc_0060F973: call [0040104Ch] ; __vbaFreeObjList
  loc_0060F979: mov eax, var_C4
  loc_0060F97F: add esp, 00000014h
  loc_0060F982: lea ecx, var_24
  loc_0060F985: mov edx, [eax]
  loc_0060F987: push ecx
  loc_0060F988: push eax
  loc_0060F989: call [edx+0000008Ch]
  loc_0060F98F: test eax, eax
  loc_0060F991: fnclex
  loc_0060F993: jge 0060F9A9h
  loc_0060F995: mov edx, var_C4
  loc_0060F99B: push 0000008Ch
  loc_0060F9A0: push 0043B5D0h
  loc_0060F9A5: push edx
  loc_0060F9A6: push eax
  loc_0060F9A7: call edi
  loc_0060F9A9: lea ebx, var_28
  loc_0060F9AC: mov eax, var_24
  loc_0060F9AF: push ebx
  loc_0060F9B0: mov edx, 00000008h
  loc_0060F9B5: sub esp, 00000010h
  loc_0060F9B8: mov var_70, edx
  loc_0060F9BB: mov ebx, esp
  loc_0060F9BD: mov ecx, 0043B848h ; "Section4"
  loc_0060F9C2: mov var_68, ecx
  loc_0060F9C5: mov esi, [eax]
  loc_0060F9C7: mov [ebx], edx
  loc_0060F9C9: mov edx, var_6C
  loc_0060F9CC: push eax
  loc_0060F9CD: mov var_AC, eax
  loc_0060F9D3: mov [ebx+00000004h], edx
  loc_0060F9D6: mov [ebx+00000008h], ecx
  loc_0060F9D9: mov ecx, var_64
  loc_0060F9DC: mov [ebx+0000000Ch], ecx
  loc_0060F9DF: call [esi+0000001Ch]
  loc_0060F9E2: test eax, eax
  loc_0060F9E4: fnclex
  loc_0060F9E6: jge 0060F9F9h
  loc_0060F9E8: mov edx, var_AC
  loc_0060F9EE: push 0000001Ch
  loc_0060F9F0: push 00439DD8h
  loc_0060F9F5: push edx
  loc_0060F9F6: push eax
  loc_0060F9F7: call edi
  loc_0060F9F9: mov eax, var_28
  loc_0060F9FC: lea edx, var_2C
  loc_0060F9FF: push edx
  loc_0060FA00: push eax
  loc_0060FA01: mov ecx, [eax]
  loc_0060FA03: mov esi, eax
  loc_0060FA05: call [ecx+0000003Ch]
  loc_0060FA08: test eax, eax
  loc_0060FA0A: fnclex
  loc_0060FA0C: jge 0060FA19h
  loc_0060FA0E: push 0000003Ch
  loc_0060FA10: push 00439BF8h
  loc_0060FA15: push esi
  loc_0060FA16: push eax
  loc_0060FA17: call edi
  loc_0060FA19: lea ebx, var_30
  loc_0060FA1C: mov eax, var_2C
  loc_0060FA1F: push ebx
  loc_0060FA20: mov edx, 00000008h
  loc_0060FA25: sub esp, 00000010h
  loc_0060FA28: mov esi, [eax]
  loc_0060FA2A: mov ebx, esp
  loc_0060FA2C: mov ecx, 0043B8A4h ; "LabelKota"
  loc_0060FA31: push eax
  loc_0060FA32: mov var_BC, eax
  loc_0060FA38: mov [ebx], edx
  loc_0060FA3A: mov edx, var_7C
  loc_0060FA3D: mov [ebx+00000004h], edx
  loc_0060FA40: mov [ebx+00000008h], ecx
  loc_0060FA43: mov ecx, var_74
  loc_0060FA46: mov [ebx+0000000Ch], ecx
  loc_0060FA49: call [esi+0000001Ch]
  loc_0060FA4C: test eax, eax
  loc_0060FA4E: fnclex
  loc_0060FA50: jge 0060FA63h
  loc_0060FA52: mov edx, var_BC
  loc_0060FA58: push 0000001Ch
  loc_0060FA5A: push 0043B7D8h
  loc_0060FA5F: push edx
  loc_0060FA60: push eax
  loc_0060FA61: call edi
  loc_0060FA63: mov ecx, [006680F8h]
  loc_0060FA69: mov edx, [006680FCh]
  loc_0060FA6F: sub esp, 00000010h
  loc_0060FA72: mov eax, esp
  loc_0060FA74: push 0043B85Ch ; "Caption"
  loc_0060FA79: mov [eax], ecx
  loc_0060FA7B: mov ecx, [00668100h]
  loc_0060FA81: mov [eax+00000004h], edx
  loc_0060FA84: mov edx, [00668104h]
  loc_0060FA8A: mov [eax+00000008h], ecx
  loc_0060FA8D: mov [eax+0000000Ch], edx
  loc_0060FA90: mov eax, var_30
  loc_0060FA93: push eax
  loc_0060FA94: call [00401094h] ; __vbaLateMemSt
  loc_0060FA9A: lea ecx, var_30
  loc_0060FA9D: lea edx, var_2C
  loc_0060FAA0: push ecx
  loc_0060FAA1: lea eax, var_28
  loc_0060FAA4: push edx
  loc_0060FAA5: lea ecx, var_24
  loc_0060FAA8: push eax
  loc_0060FAA9: push ecx
  loc_0060FAAA: push 00000004h
  loc_0060FAAC: call [0040104Ch] ; __vbaFreeObjList
  loc_0060FAB2: mov eax, var_C4
  loc_0060FAB8: add esp, 00000014h
  loc_0060FABB: lea ecx, var_24
  loc_0060FABE: mov edx, [eax]
  loc_0060FAC0: push ecx
  loc_0060FAC1: push eax
  loc_0060FAC2: call [edx+0000008Ch]
  loc_0060FAC8: test eax, eax
  loc_0060FACA: fnclex
  loc_0060FACC: jge 0060FAE2h
  loc_0060FACE: mov edx, var_C4
  loc_0060FAD4: push 0000008Ch
  loc_0060FAD9: push 0043B5D0h
  loc_0060FADE: push edx
  loc_0060FADF: push eax
  loc_0060FAE0: call edi
  loc_0060FAE2: lea ebx, var_28
  loc_0060FAE5: mov eax, var_24
  loc_0060FAE8: push ebx
  loc_0060FAE9: mov edx, 00000008h
  loc_0060FAEE: sub esp, 00000010h
  loc_0060FAF1: mov var_70, edx
  loc_0060FAF4: mov ebx, esp
  loc_0060FAF6: mov ecx, 0043B848h ; "Section4"
  loc_0060FAFB: mov var_68, ecx
  loc_0060FAFE: mov esi, [eax]
  loc_0060FB00: mov [ebx], edx
  loc_0060FB02: mov edx, var_6C
  loc_0060FB05: push eax
  loc_0060FB06: mov var_AC, eax
  loc_0060FB0C: mov [ebx+00000004h], edx
  loc_0060FB0F: mov [ebx+00000008h], ecx
  loc_0060FB12: mov ecx, var_64
  loc_0060FB15: mov [ebx+0000000Ch], ecx
  loc_0060FB18: call [esi+0000001Ch]
  loc_0060FB1B: test eax, eax
  loc_0060FB1D: fnclex
  loc_0060FB1F: jge 0060FB32h
  loc_0060FB21: mov edx, var_AC
  loc_0060FB27: push 0000001Ch
  loc_0060FB29: push 00439DD8h
  loc_0060FB2E: push edx
  loc_0060FB2F: push eax
  loc_0060FB30: call edi
  loc_0060FB32: mov eax, var_28
  loc_0060FB35: lea edx, var_2C
  loc_0060FB38: push edx
  loc_0060FB39: push eax
  loc_0060FB3A: mov ecx, [eax]
  loc_0060FB3C: mov esi, eax
  loc_0060FB3E: call [ecx+0000003Ch]
  loc_0060FB41: test eax, eax
  loc_0060FB43: fnclex
  loc_0060FB45: jge 0060FB52h
  loc_0060FB47: push 0000003Ch
  loc_0060FB49: push 00439BF8h
  loc_0060FB4E: push esi
  loc_0060FB4F: push eax
  loc_0060FB50: call edi
  loc_0060FB52: lea ebx, var_30
  loc_0060FB55: mov eax, var_2C
  loc_0060FB58: push ebx
  loc_0060FB59: mov edx, 00000008h
  loc_0060FB5E: sub esp, 00000010h
  loc_0060FB61: mov esi, [eax]
  loc_0060FB63: mov ebx, esp
  loc_0060FB65: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0060FB6A: push eax
  loc_0060FB6B: mov var_BC, eax
  loc_0060FB71: mov [ebx], edx
  loc_0060FB73: mov edx, var_7C
  loc_0060FB76: mov [ebx+00000004h], edx
  loc_0060FB79: mov [ebx+00000008h], ecx
  loc_0060FB7C: mov ecx, var_74
  loc_0060FB7F: mov [ebx+0000000Ch], ecx
  loc_0060FB82: call [esi+0000001Ch]
  loc_0060FB85: test eax, eax
  loc_0060FB87: fnclex
  loc_0060FB89: jge 0060FB9Ch
  loc_0060FB8B: mov edx, var_BC
  loc_0060FB91: push 0000001Ch
  loc_0060FB93: push 0043B7D8h
  loc_0060FB98: push edx
  loc_0060FB99: push eax
  loc_0060FB9A: call edi
  loc_0060FB9C: mov ecx, [00668108h]
  loc_0060FBA2: mov edx, [0066810Ch]
  loc_0060FBA8: sub esp, 00000010h
  loc_0060FBAB: mov eax, esp
  loc_0060FBAD: push 0043B85Ch ; "Caption"
  loc_0060FBB2: mov [eax], ecx
  loc_0060FBB4: mov ecx, [00668110h]
  loc_0060FBBA: mov [eax+00000004h], edx
  loc_0060FBBD: mov edx, [00668114h]
  loc_0060FBC3: mov [eax+00000008h], ecx
  loc_0060FBC6: mov [eax+0000000Ch], edx
  loc_0060FBC9: mov eax, var_30
  loc_0060FBCC: push eax
  loc_0060FBCD: call [00401094h] ; __vbaLateMemSt
  loc_0060FBD3: lea ecx, var_30
  loc_0060FBD6: lea edx, var_2C
  loc_0060FBD9: push ecx
  loc_0060FBDA: lea eax, var_28
  loc_0060FBDD: push edx
  loc_0060FBDE: lea ecx, var_24
  loc_0060FBE1: push eax
  loc_0060FBE2: push ecx
  loc_0060FBE3: push 00000004h
  loc_0060FBE5: call [0040104Ch] ; __vbaFreeObjList
  loc_0060FBEB: mov eax, var_C4
  loc_0060FBF1: add esp, 00000014h
  loc_0060FBF4: mov edx, [eax]
  loc_0060FBF6: push 0000000Ah
  loc_0060FBF8: push eax
  loc_0060FBF9: call [edx+00000048h]
  loc_0060FBFC: test eax, eax
  loc_0060FBFE: fnclex
  loc_0060FC00: jge 0060FC13h
  loc_0060FC02: mov ecx, var_C4
  loc_0060FC08: push 00000048h
  loc_0060FC0A: push 0043B5D0h
  loc_0060FC0F: push ecx
  loc_0060FC10: push eax
  loc_0060FC11: call edi
  loc_0060FC13: mov eax, var_C4
  loc_0060FC19: push 0000000Ah
  loc_0060FC1B: push eax
  loc_0060FC1C: mov edx, [eax]
  loc_0060FC1E: call [edx+00000058h]
  loc_0060FC21: test eax, eax
  loc_0060FC23: fnclex
  loc_0060FC25: jge 0060FC38h
  loc_0060FC27: mov ecx, var_C4
  loc_0060FC2D: push 00000058h
  loc_0060FC2F: push 0043B5D0h
  loc_0060FC34: push ecx
  loc_0060FC35: push eax
  loc_0060FC36: call edi
  loc_0060FC38: sub esp, 00000010h
  loc_0060FC3B: mov eax, 00000002h
  loc_0060FC40: mov edx, esp
  loc_0060FC42: mov ecx, eax
  loc_0060FC44: mov var_70, ecx
  loc_0060FC47: mov var_68, eax
  loc_0060FC4A: mov [edx], ecx
  loc_0060FC4C: mov ecx, var_6C
  loc_0060FC4F: push 8001000Eh
  loc_0060FC54: mov [edx+00000004h], ecx
  loc_0060FC57: mov ecx, var_C4
  loc_0060FC5D: push ecx
  loc_0060FC5E: mov [edx+00000008h], eax
  loc_0060FC61: mov eax, var_64
  loc_0060FC64: mov [edx+0000000Ch], eax
  loc_0060FC67: call [00401280h] ; __vbaLateIdSt
  loc_0060FC6D: mov edx, var_C4
  loc_0060FC73: push 00000000h
  loc_0060FC75: push 80011003h
  loc_0060FC7A: push edx
  loc_0060FC7B: call [0040102Ch] ; __vbaLateIdCall
  loc_0060FC81: add esp, 0000000Ch
  loc_0060FC84: lea eax, var_C4
  loc_0060FC8A: push 00000000h
  loc_0060FC8C: push eax
  loc_0060FC8D: call [004010C8h] ; __vbaObjSetAddref
  loc_0060FC93: push 0060FCFDh
  loc_0060FC98: jmp 0060FCD3h
  loc_0060FC9A: lea ecx, var_20
  loc_0060FC9D: call [0040129Ch] ; __vbaFreeStr
  loc_0060FCA3: lea ecx, var_30
  loc_0060FCA6: lea edx, var_2C
  loc_0060FCA9: push ecx
  loc_0060FCAA: lea eax, var_28
  loc_0060FCAD: push edx
  loc_0060FCAE: lea ecx, var_24
  loc_0060FCB1: push eax
  loc_0060FCB2: push ecx
  loc_0060FCB3: push 00000004h
  loc_0060FCB5: call [0040104Ch] ; __vbaFreeObjList
  loc_0060FCBB: lea edx, var_60
  loc_0060FCBE: lea eax, var_50
  loc_0060FCC1: push edx
  loc_0060FCC2: lea ecx, var_40
  loc_0060FCC5: push eax
  loc_0060FCC6: push ecx
  loc_0060FCC7: push 00000003h
  loc_0060FCC9: call [0040103Ch] ; __vbaFreeVarList
  loc_0060FCCF: add esp, 00000024h
  loc_0060FCD2: ret
  loc_0060FCD3: lea edx, var_C8
  loc_0060FCD9: lea eax, var_C4
  loc_0060FCDF: push edx
  loc_0060FCE0: push eax
  loc_0060FCE1: push 00000002h
  loc_0060FCE3: call [0040104Ch] ; __vbaFreeObjList
  loc_0060FCE9: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0060FCEF: add esp, 0000000Ch
  loc_0060FCF2: lea ecx, var_18
  loc_0060FCF5: call __vbaFreeStr
  loc_0060FCF7: lea ecx, var_1C
  loc_0060FCFA: call __vbaFreeStr
  loc_0060FCFC: ret
  loc_0060FCFD: mov ecx, var_10
  loc_0060FD00: pop edi
  loc_0060FD01: pop esi
  loc_0060FD02: mov fs:[00000000h], ecx
  loc_0060FD09: pop ebx
  loc_0060FD0A: mov esp, ebp
  loc_0060FD0C: pop ebp
  loc_0060FD0D: ret
End Sub

Private Sub Proc_34_7_60FD20
  loc_0060FD20: push ebp
  loc_0060FD21: mov ebp, esp
  loc_0060FD23: sub esp, 00000008h
  loc_0060FD26: push 00405696h ; __vbaExceptHandler
  loc_0060FD2B: mov eax, fs:[00000000h]
  loc_0060FD31: push eax
  loc_0060FD32: mov fs:[00000000h], esp
  loc_0060FD39: sub esp, 000000D0h
  loc_0060FD3F: push ebx
  loc_0060FD40: push esi
  loc_0060FD41: push edi
  loc_0060FD42: mov var_8, esp
  loc_0060FD45: mov var_4, 00404510h
  loc_0060FD4C: xor ebx, ebx
  loc_0060FD4E: mov var_18, ebx
  loc_0060FD51: mov var_1C, ebx
  loc_0060FD54: mov var_20, ebx
  loc_0060FD57: mov var_24, ebx
  loc_0060FD5A: mov var_28, ebx
  loc_0060FD5D: mov var_2C, ebx
  loc_0060FD60: mov var_30, ebx
  loc_0060FD63: mov var_40, ebx
  loc_0060FD66: mov var_50, ebx
  loc_0060FD69: mov var_60, ebx
  loc_0060FD6C: mov var_70, ebx
  loc_0060FD6F: mov var_80, ebx
  loc_0060FD72: mov var_90, ebx
  loc_0060FD78: mov var_A4, ebx
  loc_0060FD7E: mov var_C4, ebx
  loc_0060FD84: mov var_C8, ebx
  loc_0060FD8A: call 0055B320h
  loc_0060FD8F: mov esi, [004011FCh] ; __vbaStrCopy
  loc_0060FD95: mov edx, 0042DFECh
  loc_0060FD9A: lea ecx, var_18
  loc_0060FD9D: call __vbaStrCopy
  loc_0060FD9F: mov edx, 0043BF0Ch ; " SELECT no_nota, tgl_nota, nama_pelanggan, jenis, tot_jual from penjualan order by no_nota"
  loc_0060FDA4: lea ecx, var_18
  loc_0060FDA7: call __vbaStrCopy
  loc_0060FDA9: push 0042DF30h
  loc_0060FDAE: call [00401168h] ; __vbaNew
  loc_0060FDB4: push eax
  loc_0060FDB5: push 00668090h
  loc_0060FDBA: call [004010B8h] ; __vbaObjSet
  loc_0060FDC0: cmp [0066803Ch], ebx
  loc_0060FDC6: jnz 0060FDD8h
  loc_0060FDC8: push 0066803Ch
  loc_0060FDCD: push 0042DEFCh
  loc_0060FDD2: call [004011E8h] ; __vbaNew2
  loc_0060FDD8: mov esi, [0066803Ch]
  loc_0060FDDE: lea ecx, var_40
  loc_0060FDE1: mov var_38, 80020004h
  loc_0060FDE8: mov var_40, 0000000Ah
  loc_0060FDEF: call [0040123Ch] ; __vbaFreeVarg
  loc_0060FDF5: mov eax, [esi]
  loc_0060FDF7: lea ecx, var_24
  loc_0060FDFA: push ecx
  loc_0060FDFB: mov ecx, var_18
  loc_0060FDFE: lea edx, var_40
  loc_0060FE01: push 00000001h
  loc_0060FE03: push edx
  loc_0060FE04: push ecx
  loc_0060FE05: push esi
  loc_0060FE06: call [eax+00000040h]
  loc_0060FE09: cmp eax, ebx
  loc_0060FE0B: fnclex
  loc_0060FE0D: jge 0060FE22h
  loc_0060FE0F: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060FE15: push 00000040h
  loc_0060FE17: push 0042E1B0h
  loc_0060FE1C: push esi
  loc_0060FE1D: push eax
  loc_0060FE1E: call edi
  loc_0060FE20: jmp 0060FE28h
  loc_0060FE22: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0060FE28: mov edx, var_24
  loc_0060FE2B: mov esi, [00401268h] ; __vbaCastObj
  loc_0060FE31: push 0042DF20h
  loc_0060FE36: push edx
  loc_0060FE37: call __vbaCastObj
  loc_0060FE39: push eax
  loc_0060FE3A: push 00668090h
  loc_0060FE3F: call [004010B8h] ; __vbaObjSet
  loc_0060FE45: lea ecx, var_24
  loc_0060FE48: call [004012A0h] ; __vbaFreeObj
  loc_0060FE4E: lea ecx, var_40
  loc_0060FE51: call [00401020h] ; __vbaFreeVar
  loc_0060FE57: cmp [00668580h], ebx
  loc_0060FE5D: jnz 0060FE6Fh
  loc_0060FE5F: push 00668580h
  loc_0060FE64: push 00406F48h
  loc_0060FE69: call [004011E8h] ; __vbaNew2
  loc_0060FE6F: mov eax, [00668580h]
  loc_0060FE74: lea ecx, var_C4
  loc_0060FE7A: push eax
  loc_0060FE7B: push ecx
  loc_0060FE7C: call [004010C8h] ; __vbaObjSetAddref
  loc_0060FE82: mov edx, var_C4
  loc_0060FE88: push 0043B7B8h
  loc_0060FE8D: push ebx
  loc_0060FE8E: mov edx, [edx]
  loc_0060FE90: mov var_E0, edx
  loc_0060FE96: call __vbaCastObj
  loc_0060FE98: push eax
  loc_0060FE99: lea eax, var_24
  loc_0060FE9C: push eax
  loc_0060FE9D: call [004010B8h] ; __vbaObjSet
  loc_0060FEA3: mov ecx, var_C4
  loc_0060FEA9: mov edx, var_E0
  loc_0060FEAF: push eax
  loc_0060FEB0: push ecx
  loc_0060FEB1: call [edx+00000028h]
  loc_0060FEB4: cmp eax, ebx
  loc_0060FEB6: fnclex
  loc_0060FEB8: jge 0060FECBh
  loc_0060FEBA: mov ecx, var_C4
  loc_0060FEC0: push 00000028h
  loc_0060FEC2: push 0043B5D0h
  loc_0060FEC7: push ecx
  loc_0060FEC8: push eax
  loc_0060FEC9: call edi
  loc_0060FECB: lea ecx, var_24
  loc_0060FECE: call [004012A0h] ; __vbaFreeObj
  loc_0060FED4: mov eax, var_C4
  loc_0060FEDA: push 0042DFECh
  loc_0060FEDF: push eax
  loc_0060FEE0: mov edx, [eax]
  loc_0060FEE2: call [edx+00000030h]
  loc_0060FEE5: cmp eax, ebx
  loc_0060FEE7: fnclex
  loc_0060FEE9: jge 0060FEFCh
  loc_0060FEEB: mov ecx, var_C4
  loc_0060FEF1: push 00000030h
  loc_0060FEF3: push 0043B5D0h
  loc_0060FEF8: push ecx
  loc_0060FEF9: push eax
  loc_0060FEFA: call edi
  loc_0060FEFC: mov eax, [00668090h]
  loc_0060FF01: lea ecx, var_24
  loc_0060FF04: push ecx
  loc_0060FF05: push eax
  loc_0060FF06: mov edx, [eax]
  loc_0060FF08: call [edx+00000114h]
  loc_0060FF0E: cmp eax, ebx
  loc_0060FF10: fnclex
  loc_0060FF12: jge 0060FF28h
  loc_0060FF14: mov edx, [00668090h]
  loc_0060FF1A: push 00000114h
  loc_0060FF1F: push 0043B7C8h
  loc_0060FF24: push edx
  loc_0060FF25: push eax
  loc_0060FF26: call edi
  loc_0060FF28: mov eax, var_C4
  loc_0060FF2E: mov ecx, var_24
  loc_0060FF31: push 0043B7B8h
  loc_0060FF36: push ecx
  loc_0060FF37: mov esi, [eax]
  loc_0060FF39: call [00401268h] ; __vbaCastObj
  loc_0060FF3F: lea edx, var_28
  loc_0060FF42: push eax
  loc_0060FF43: push edx
  loc_0060FF44: call [004010B8h] ; __vbaObjSet
  loc_0060FF4A: push eax
  loc_0060FF4B: mov eax, var_C4
  loc_0060FF51: push eax
  loc_0060FF52: call [esi+00000028h]
  loc_0060FF55: cmp eax, ebx
  loc_0060FF57: fnclex
  loc_0060FF59: jge 0060FF6Ch
  loc_0060FF5B: mov ecx, var_C4
  loc_0060FF61: push 00000028h
  loc_0060FF63: push 0043B5D0h
  loc_0060FF68: push ecx
  loc_0060FF69: push eax
  loc_0060FF6A: call edi
  loc_0060FF6C: lea edx, var_28
  loc_0060FF6F: lea eax, var_24
  loc_0060FF72: push edx
  loc_0060FF73: push eax
  loc_0060FF74: push 00000002h
  loc_0060FF76: call [0040104Ch] ; __vbaFreeObjList
  loc_0060FF7C: mov eax, var_C4
  loc_0060FF82: add esp, 0000000Ch
  loc_0060FF85: lea edx, var_24
  loc_0060FF88: mov ecx, [eax]
  loc_0060FF8A: push edx
  loc_0060FF8B: push eax
  loc_0060FF8C: call [ecx+0000008Ch]
  loc_0060FF92: cmp eax, ebx
  loc_0060FF94: fnclex
  loc_0060FF96: jge 0060FFACh
  loc_0060FF98: mov ecx, var_C4
  loc_0060FF9E: push 0000008Ch
  loc_0060FFA3: push 0043B5D0h
  loc_0060FFA8: push ecx
  loc_0060FFA9: push eax
  loc_0060FFAA: call edi
  loc_0060FFAC: lea esi, var_28
  loc_0060FFAF: mov eax, var_24
  loc_0060FFB2: push esi
  loc_0060FFB3: mov ecx, 00000008h
  loc_0060FFB8: sub esp, 00000010h
  loc_0060FFBB: mov var_70, ecx
  loc_0060FFBE: mov esi, esp
  loc_0060FFC0: mov var_68, 00439FB0h ; "Section1"
  loc_0060FFC7: mov edx, [eax]
  loc_0060FFC9: push eax
  loc_0060FFCA: mov [esi], ecx
  loc_0060FFCC: mov ecx, var_6C
  loc_0060FFCF: mov var_AC, eax
  loc_0060FFD5: mov [esi+00000004h], ecx
  loc_0060FFD8: mov ecx, var_68
  loc_0060FFDB: mov [esi+00000008h], ecx
  loc_0060FFDE: mov ecx, var_64
  loc_0060FFE1: mov [esi+0000000Ch], ecx
  loc_0060FFE4: call [edx+0000001Ch]
  loc_0060FFE7: cmp eax, ebx
  loc_0060FFE9: fnclex
  loc_0060FFEB: jge 0060FFFEh
  loc_0060FFED: mov edx, var_AC
  loc_0060FFF3: push 0000001Ch
  loc_0060FFF5: push 00439DD8h
  loc_0060FFFA: push edx
  loc_0060FFFB: push eax
  loc_0060FFFC: call edi
  loc_0060FFFE: mov eax, var_28
  loc_00610001: lea edx, var_2C
  loc_00610004: push edx
  loc_00610005: push eax
  loc_00610006: mov ecx, [eax]
  loc_00610008: mov esi, eax
  loc_0061000A: call [ecx+0000003Ch]
  loc_0061000D: cmp eax, ebx
  loc_0061000F: fnclex
  loc_00610011: jge 0061001Eh
  loc_00610013: push 0000003Ch
  loc_00610015: push 00439BF8h
  loc_0061001A: push esi
  loc_0061001B: push eax
  loc_0061001C: call edi
  loc_0061001E: mov eax, var_2C
  loc_00610021: mov var_2C, ebx
  loc_00610024: push eax
  loc_00610025: lea eax, var_C8
  loc_0061002B: push eax
  loc_0061002C: call [004010B8h] ; __vbaObjSet
  loc_00610032: lea ecx, var_28
  loc_00610035: lea edx, var_24
  loc_00610038: push ecx
  loc_00610039: push edx
  loc_0061003A: push 00000002h
  loc_0061003C: call [0040104Ch] ; __vbaFreeObjList
  loc_00610042: mov eax, var_C8
  loc_00610048: add esp, 0000000Ch
  loc_0061004B: lea edx, var_A4
  loc_00610051: mov ecx, [eax]
  loc_00610053: push edx
  loc_00610054: push eax
  loc_00610055: call [ecx+00000020h]
  loc_00610058: cmp eax, ebx
  loc_0061005A: fnclex
  loc_0061005C: jge 0061006Fh
  loc_0061005E: mov ecx, var_C8
  loc_00610064: push 00000020h
  loc_00610066: push 0043B7D8h
  loc_0061006B: push ecx
  loc_0061006C: push eax
  loc_0061006D: call edi
  loc_0061006F: mov ecx, var_A4
  loc_00610075: call [00401138h] ; __vbaI2I4
  loc_0061007B: mov var_D0, eax
  loc_00610081: mov ebx, 00000001h
  loc_00610086: cmp bx, var_D0
  loc_0061008D: mov var_14, ebx
  loc_00610090: jg 00610301h
  loc_00610096: lea esi, var_24
  loc_00610099: mov ecx, var_C8
  loc_0061009F: push esi
  loc_006100A0: mov eax, 00000002h
  loc_006100A5: sub esp, 00000010h
  loc_006100A8: mov var_70, eax
  loc_006100AB: mov esi, esp
  loc_006100AD: mov var_68, bx
  loc_006100B1: mov edx, [ecx]
  loc_006100B3: push ecx
  loc_006100B4: mov [esi], eax
  loc_006100B6: mov eax, var_6C
  loc_006100B9: mov [esi+00000004h], eax
  loc_006100BC: mov eax, var_68
  loc_006100BF: mov [esi+00000008h], eax
  loc_006100C2: mov eax, var_64
  loc_006100C5: mov [esi+0000000Ch], eax
  loc_006100C8: call [edx+0000001Ch]
  loc_006100CB: test eax, eax
  loc_006100CD: fnclex
  loc_006100CF: jge 006100E2h
  loc_006100D1: mov ecx, var_C8
  loc_006100D7: push 0000001Ch
  loc_006100D9: push 0043B7D8h
  loc_006100DE: push ecx
  loc_006100DF: push eax
  loc_006100E0: call edi
  loc_006100E2: mov edx, var_24
  loc_006100E5: push 0043B7E8h
  loc_006100EA: push edx
  loc_006100EB: call [004011CCh] ; __vbaCheckType
  loc_006100F1: lea ecx, var_24
  loc_006100F4: mov si, ax
  loc_006100F7: call [004012A0h] ; __vbaFreeObj
  loc_006100FD: test si, si
  loc_00610100: Unknown_795()
  loc_00610106: mov var_68, bx
  loc_0061010A: lea ebx, var_24
  loc_0061010D: push ebx
  loc_0061010E: mov ecx, var_C8
  loc_00610114: sub esp, 00000010h
  loc_00610117: mov eax, 00000002h
  loc_0061011C: mov ebx, esp
  loc_0061011E: mov var_70, eax
  loc_00610121: mov edx, [ecx]
  loc_00610123: mov esi, 0042DFECh
  loc_00610128: mov [ebx], eax
  loc_0061012A: mov eax, var_6C
  loc_0061012D: push ecx
  loc_0061012E: mov var_78, esi
  loc_00610131: mov [ebx+00000004h], eax
  loc_00610134: mov eax, var_68
  loc_00610137: mov edi, 00000008h
  loc_0061013C: mov [ebx+00000008h], eax
  loc_0061013F: mov eax, var_64
  loc_00610142: mov [ebx+0000000Ch], eax
  loc_00610145: call [edx+0000001Ch]
  loc_00610148: test eax, eax
  loc_0061014A: fnclex
  loc_0061014C: jge 00610163h
  loc_0061014E: mov ecx, var_C8
  loc_00610154: push 0000001Ch
  loc_00610156: push 0043B7D8h
  loc_0061015B: push ecx
  loc_0061015C: push eax
  loc_0061015D: call [00401080h] ; __vbaHresultCheckObj
  loc_00610163: mov eax, var_7C
  loc_00610166: sub esp, 00000010h
  loc_00610169: mov edx, esp
  loc_0061016B: mov ecx, var_74
  loc_0061016E: push 0043B7F8h ; "DataMember"
  loc_00610173: mov [edx], edi
  loc_00610175: mov [edx+00000004h], eax
  loc_00610178: mov [edx+00000008h], esi
  loc_0061017B: mov [edx+0000000Ch], ecx
  loc_0061017E: mov edx, var_24
  loc_00610181: push edx
  loc_00610182: call [00401094h] ; __vbaLateMemSt
  loc_00610188: lea ecx, var_24
  loc_0061018B: call [004012A0h] ; __vbaFreeObj
  loc_00610191: mov eax, [00668090h]
  loc_00610196: lea edx, var_24
  loc_00610199: push edx
  loc_0061019A: push eax
  loc_0061019B: mov ecx, [eax]
  loc_0061019D: call [ecx+00000054h]
  loc_006101A0: test eax, eax
  loc_006101A2: fnclex
  loc_006101A4: jge 006101BBh
  loc_006101A6: mov ecx, [00668090h]
  loc_006101AC: push 00000054h
  loc_006101AE: push 0042DF44h
  loc_006101B3: push ecx
  loc_006101B4: push eax
  loc_006101B5: call [00401080h] ; __vbaHresultCheckObj
  loc_006101BB: mov ebx, var_14
  loc_006101BE: lea edi, var_28
  loc_006101C1: mov dx, bx
  loc_006101C4: push edi
  loc_006101C5: sub dx, 0001h
  loc_006101C9: mov eax, var_24
  loc_006101CC: jo 00610A8Fh
  loc_006101D2: sub esp, 00000010h
  loc_006101D5: mov ecx, 00000002h
  loc_006101DA: mov edi, esp
  loc_006101DC: mov var_70, ecx
  loc_006101DF: mov var_68, dx
  loc_006101E3: mov edx, [eax]
  loc_006101E5: mov [edi], ecx
  loc_006101E7: mov ecx, var_6C
  loc_006101EA: push eax
  loc_006101EB: mov esi, eax
  loc_006101ED: mov [edi+00000004h], ecx
  loc_006101F0: mov ecx, var_68
  loc_006101F3: mov [edi+00000008h], ecx
  loc_006101F6: mov ecx, var_64
  loc_006101F9: mov [edi+0000000Ch], ecx
  loc_006101FC: call [edx+00000028h]
  loc_006101FF: test eax, eax
  loc_00610201: fnclex
  loc_00610203: jge 00610218h
  loc_00610205: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0061020B: push 00000028h
  loc_0061020D: push 0042DFACh
  loc_00610212: push esi
  loc_00610213: push eax
  loc_00610214: call edi
  loc_00610216: jmp 0061021Eh
  loc_00610218: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0061021E: mov eax, var_28
  loc_00610221: lea ecx, var_20
  loc_00610224: push ecx
  loc_00610225: push eax
  loc_00610226: mov edx, [eax]
  loc_00610228: mov esi, eax
  loc_0061022A: call [edx+0000002Ch]
  loc_0061022D: test eax, eax
  loc_0061022F: fnclex
  loc_00610231: jge 0061023Eh
  loc_00610233: push 0000002Ch
  loc_00610235: push 0042DFBCh
  loc_0061023A: push esi
  loc_0061023B: push eax
  loc_0061023C: call edi
  loc_0061023E: mov eax, var_20
  loc_00610241: lea esi, var_2C
  loc_00610244: push esi
  loc_00610245: mov ecx, var_C8
  loc_0061024B: sub esp, 00000010h
  loc_0061024E: mov var_38, eax
  loc_00610251: mov esi, esp
  loc_00610253: mov eax, 00000002h
  loc_00610258: mov var_78, bx
  loc_0061025C: mov var_20, 00000000h
  loc_00610263: mov [esi], eax
  loc_00610265: mov eax, var_7C
  loc_00610268: mov var_40, 00000008h
  loc_0061026F: mov edx, [ecx]
  loc_00610271: mov [esi+00000004h], eax
  loc_00610274: mov eax, var_78
  loc_00610277: push ecx
  loc_00610278: mov [esi+00000008h], eax
  loc_0061027B: mov eax, var_74
  loc_0061027E: mov [esi+0000000Ch], eax
  loc_00610281: call [edx+0000001Ch]
  loc_00610284: test eax, eax
  loc_00610286: fnclex
  loc_00610288: jge 0061029Bh
  loc_0061028A: mov ecx, var_C8
  loc_00610290: push 0000001Ch
  loc_00610292: push 0043B7D8h
  loc_00610297: push ecx
  loc_00610298: push eax
  loc_00610299: call edi
  loc_0061029B: mov eax, var_40
  loc_0061029E: mov ecx, var_3C
  loc_006102A1: sub esp, 00000010h
  loc_006102A4: mov edx, esp
  loc_006102A6: push 0043B810h ; "DataField"
  loc_006102AB: mov [edx], eax
  loc_006102AD: mov eax, var_38
  loc_006102B0: mov [edx+00000004h], ecx
  loc_006102B3: mov ecx, var_34
  loc_006102B6: mov [edx+00000008h], eax
  loc_006102B9: mov [edx+0000000Ch], ecx
  loc_006102BC: mov edx, var_2C
  loc_006102BF: push edx
  loc_006102C0: call [00401094h] ; __vbaLateMemSt
  loc_006102C6: lea eax, var_2C
  loc_006102C9: lea ecx, var_28
  loc_006102CC: push eax
  loc_006102CD: lea edx, var_24
  loc_006102D0: push ecx
  loc_006102D1: push edx
  loc_006102D2: push 00000003h
  loc_006102D4: call [0040104Ch] ; __vbaFreeObjList
  loc_006102DA: add esp, 00000010h
  loc_006102DD: lea ecx, var_40
  loc_006102E0: call [00401020h] ; __vbaFreeVar
  loc_006102E6: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_006102EC: mov eax, 00000001h
  loc_006102F1: add ax, bx
  loc_006102F4: jo 00610A8Fh
  loc_006102FA: mov ebx, eax
  loc_006102FC: jmp 00610086h
  loc_00610301: lea eax, var_C8
  loc_00610307: push 00000000h
  loc_00610309: push eax
  loc_0061030A: call [004010C8h] ; __vbaObjSetAddref
  loc_00610310: lea ecx, var_40
  loc_00610313: push ecx
  loc_00610314: call [00401220h] ; rtcGetDateVar
  loc_0061031A: lea edx, var_70
  loc_0061031D: lea ecx, var_50
  loc_00610320: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_00610327: mov var_70, 00000008h
  loc_0061032E: call [00401238h] ; __vbaVarDup
  loc_00610334: push 00000001h
  loc_00610336: lea edx, var_50
  loc_00610339: push 00000001h
  loc_0061033B: lea eax, var_40
  loc_0061033E: push edx
  loc_0061033F: lea ecx, var_60
  loc_00610342: push eax
  loc_00610343: push ecx
  loc_00610344: call [00401078h] ; rtcVarFromFormatVar
  loc_0061034A: mov eax, var_C4
  loc_00610350: lea ecx, var_24
  loc_00610353: push ecx
  loc_00610354: push eax
  loc_00610355: mov edx, [eax]
  loc_00610357: call [edx+0000008Ch]
  loc_0061035D: test eax, eax
  loc_0061035F: fnclex
  loc_00610361: jge 00610377h
  loc_00610363: mov edx, var_C4
  loc_00610369: push 0000008Ch
  loc_0061036E: push 0043B5D0h
  loc_00610373: push edx
  loc_00610374: push eax
  loc_00610375: call edi
  loc_00610377: lea ebx, var_28
  loc_0061037A: mov eax, var_24
  loc_0061037D: push ebx
  loc_0061037E: mov edx, 00000008h
  loc_00610383: sub esp, 00000010h
  loc_00610386: mov esi, [eax]
  loc_00610388: mov ebx, esp
  loc_0061038A: mov ecx, 0043B848h ; "Section4"
  loc_0061038F: push eax
  loc_00610390: mov var_AC, eax
  loc_00610396: mov [ebx], edx
  loc_00610398: mov edx, var_7C
  loc_0061039B: mov [ebx+00000004h], edx
  loc_0061039E: mov [ebx+00000008h], ecx
  loc_006103A1: mov ecx, var_74
  loc_006103A4: mov [ebx+0000000Ch], ecx
  loc_006103A7: call [esi+0000001Ch]
  loc_006103AA: test eax, eax
  loc_006103AC: fnclex
  loc_006103AE: jge 006103C1h
  loc_006103B0: mov edx, var_AC
  loc_006103B6: push 0000001Ch
  loc_006103B8: push 00439DD8h
  loc_006103BD: push edx
  loc_006103BE: push eax
  loc_006103BF: call edi
  loc_006103C1: mov eax, var_28
  loc_006103C4: lea edx, var_2C
  loc_006103C7: push edx
  loc_006103C8: push eax
  loc_006103C9: mov ecx, [eax]
  loc_006103CB: mov esi, eax
  loc_006103CD: call [ecx+0000003Ch]
  loc_006103D0: test eax, eax
  loc_006103D2: fnclex
  loc_006103D4: jge 006103E1h
  loc_006103D6: push 0000003Ch
  loc_006103D8: push 00439BF8h
  loc_006103DD: push esi
  loc_006103DE: push eax
  loc_006103DF: call edi
  loc_006103E1: lea ebx, var_30
  loc_006103E4: mov eax, var_2C
  loc_006103E7: push ebx
  loc_006103E8: mov edx, 00000008h
  loc_006103ED: sub esp, 00000010h
  loc_006103F0: mov esi, [eax]
  loc_006103F2: mov ebx, esp
  loc_006103F4: mov ecx, 0043B828h ; "labelTanggal"
  loc_006103F9: push eax
  loc_006103FA: mov var_BC, eax
  loc_00610400: mov [ebx], edx
  loc_00610402: mov edx, var_8C
  loc_00610408: mov [ebx+00000004h], edx
  loc_0061040B: mov [ebx+00000008h], ecx
  loc_0061040E: mov ecx, var_84
  loc_00610414: mov [ebx+0000000Ch], ecx
  loc_00610417: call [esi+0000001Ch]
  loc_0061041A: test eax, eax
  loc_0061041C: fnclex
  loc_0061041E: jge 00610431h
  loc_00610420: mov edx, var_BC
  loc_00610426: push 0000001Ch
  loc_00610428: push 0043B7D8h
  loc_0061042D: push edx
  loc_0061042E: push eax
  loc_0061042F: call edi
  loc_00610431: mov ecx, var_60
  loc_00610434: mov edx, var_5C
  loc_00610437: sub esp, 00000010h
  loc_0061043A: mov eax, esp
  loc_0061043C: push 0043B85Ch ; "Caption"
  loc_00610441: mov [eax], ecx
  loc_00610443: mov ecx, var_58
  loc_00610446: mov [eax+00000004h], edx
  loc_00610449: mov edx, var_54
  loc_0061044C: mov [eax+00000008h], ecx
  loc_0061044F: mov [eax+0000000Ch], edx
  loc_00610452: mov eax, var_30
  loc_00610455: push eax
  loc_00610456: call [00401094h] ; __vbaLateMemSt
  loc_0061045C: lea ecx, var_30
  loc_0061045F: lea edx, var_2C
  loc_00610462: push ecx
  loc_00610463: lea eax, var_28
  loc_00610466: push edx
  loc_00610467: lea ecx, var_24
  loc_0061046A: push eax
  loc_0061046B: push ecx
  loc_0061046C: push 00000004h
  loc_0061046E: call [0040104Ch] ; __vbaFreeObjList
  loc_00610474: lea edx, var_60
  loc_00610477: lea eax, var_50
  loc_0061047A: push edx
  loc_0061047B: lea ecx, var_40
  loc_0061047E: push eax
  loc_0061047F: push ecx
  loc_00610480: push 00000003h
  loc_00610482: call [0040103Ch] ; __vbaFreeVarList
  loc_00610488: mov eax, var_C4
  loc_0061048E: add esp, 00000024h
  loc_00610491: lea ecx, var_24
  loc_00610494: mov edx, [eax]
  loc_00610496: push ecx
  loc_00610497: push eax
  loc_00610498: call [edx+0000008Ch]
  loc_0061049E: test eax, eax
  loc_006104A0: fnclex
  loc_006104A2: jge 006104B8h
  loc_006104A4: mov edx, var_C4
  loc_006104AA: push 0000008Ch
  loc_006104AF: push 0043B5D0h
  loc_006104B4: push edx
  loc_006104B5: push eax
  loc_006104B6: call edi
  loc_006104B8: lea ebx, var_28
  loc_006104BB: mov eax, var_24
  loc_006104BE: push ebx
  loc_006104BF: mov edx, 00000008h
  loc_006104C4: sub esp, 00000010h
  loc_006104C7: mov var_70, edx
  loc_006104CA: mov ebx, esp
  loc_006104CC: mov ecx, 0043B848h ; "Section4"
  loc_006104D1: mov var_68, ecx
  loc_006104D4: mov esi, [eax]
  loc_006104D6: mov [ebx], edx
  loc_006104D8: mov edx, var_6C
  loc_006104DB: push eax
  loc_006104DC: mov var_AC, eax
  loc_006104E2: mov [ebx+00000004h], edx
  loc_006104E5: mov [ebx+00000008h], ecx
  loc_006104E8: mov ecx, var_64
  loc_006104EB: mov [ebx+0000000Ch], ecx
  loc_006104EE: call [esi+0000001Ch]
  loc_006104F1: test eax, eax
  loc_006104F3: fnclex
  loc_006104F5: jge 00610508h
  loc_006104F7: mov edx, var_AC
  loc_006104FD: push 0000001Ch
  loc_006104FF: push 00439DD8h
  loc_00610504: push edx
  loc_00610505: push eax
  loc_00610506: call edi
  loc_00610508: mov eax, var_28
  loc_0061050B: lea edx, var_2C
  loc_0061050E: push edx
  loc_0061050F: push eax
  loc_00610510: mov ecx, [eax]
  loc_00610512: mov esi, eax
  loc_00610514: call [ecx+0000003Ch]
  loc_00610517: test eax, eax
  loc_00610519: fnclex
  loc_0061051B: jge 00610528h
  loc_0061051D: push 0000003Ch
  loc_0061051F: push 00439BF8h
  loc_00610524: push esi
  loc_00610525: push eax
  loc_00610526: call edi
  loc_00610528: lea ebx, var_30
  loc_0061052B: mov eax, var_2C
  loc_0061052E: push ebx
  loc_0061052F: mov edx, 00000008h
  loc_00610534: sub esp, 00000010h
  loc_00610537: mov esi, [eax]
  loc_00610539: mov ebx, esp
  loc_0061053B: mov ecx, 0043B870h ; "LabelNama"
  loc_00610540: push eax
  loc_00610541: mov var_BC, eax
  loc_00610547: mov [ebx], edx
  loc_00610549: mov edx, var_7C
  loc_0061054C: mov [ebx+00000004h], edx
  loc_0061054F: mov [ebx+00000008h], ecx
  loc_00610552: mov ecx, var_74
  loc_00610555: mov [ebx+0000000Ch], ecx
  loc_00610558: call [esi+0000001Ch]
  loc_0061055B: test eax, eax
  loc_0061055D: fnclex
  loc_0061055F: jge 00610572h
  loc_00610561: mov edx, var_BC
  loc_00610567: push 0000001Ch
  loc_00610569: push 0043B7D8h
  loc_0061056E: push edx
  loc_0061056F: push eax
  loc_00610570: call edi
  loc_00610572: mov ecx, [006680D8h]
  loc_00610578: mov edx, [006680DCh]
  loc_0061057E: sub esp, 00000010h
  loc_00610581: mov eax, esp
  loc_00610583: push 0043B85Ch ; "Caption"
  loc_00610588: mov [eax], ecx
  loc_0061058A: mov ecx, [006680E0h]
  loc_00610590: mov [eax+00000004h], edx
  loc_00610593: mov edx, [006680E4h]
  loc_00610599: mov [eax+00000008h], ecx
  loc_0061059C: mov [eax+0000000Ch], edx
  loc_0061059F: mov eax, var_30
  loc_006105A2: push eax
  loc_006105A3: call [00401094h] ; __vbaLateMemSt
  loc_006105A9: lea ecx, var_30
  loc_006105AC: lea edx, var_2C
  loc_006105AF: push ecx
  loc_006105B0: lea eax, var_28
  loc_006105B3: push edx
  loc_006105B4: lea ecx, var_24
  loc_006105B7: push eax
  loc_006105B8: push ecx
  loc_006105B9: push 00000004h
  loc_006105BB: call [0040104Ch] ; __vbaFreeObjList
  loc_006105C1: mov eax, var_C4
  loc_006105C7: add esp, 00000014h
  loc_006105CA: lea ecx, var_24
  loc_006105CD: mov edx, [eax]
  loc_006105CF: push ecx
  loc_006105D0: push eax
  loc_006105D1: call [edx+0000008Ch]
  loc_006105D7: test eax, eax
  loc_006105D9: fnclex
  loc_006105DB: jge 006105F1h
  loc_006105DD: mov edx, var_C4
  loc_006105E3: push 0000008Ch
  loc_006105E8: push 0043B5D0h
  loc_006105ED: push edx
  loc_006105EE: push eax
  loc_006105EF: call edi
  loc_006105F1: lea ebx, var_28
  loc_006105F4: mov eax, var_24
  loc_006105F7: push ebx
  loc_006105F8: mov edx, 00000008h
  loc_006105FD: sub esp, 00000010h
  loc_00610600: mov var_70, edx
  loc_00610603: mov ebx, esp
  loc_00610605: mov ecx, 0043B848h ; "Section4"
  loc_0061060A: mov var_68, ecx
  loc_0061060D: mov esi, [eax]
  loc_0061060F: mov [ebx], edx
  loc_00610611: mov edx, var_6C
  loc_00610614: push eax
  loc_00610615: mov var_AC, eax
  loc_0061061B: mov [ebx+00000004h], edx
  loc_0061061E: mov [ebx+00000008h], ecx
  loc_00610621: mov ecx, var_64
  loc_00610624: mov [ebx+0000000Ch], ecx
  loc_00610627: call [esi+0000001Ch]
  loc_0061062A: test eax, eax
  loc_0061062C: fnclex
  loc_0061062E: jge 00610641h
  loc_00610630: mov edx, var_AC
  loc_00610636: push 0000001Ch
  loc_00610638: push 00439DD8h
  loc_0061063D: push edx
  loc_0061063E: push eax
  loc_0061063F: call edi
  loc_00610641: mov eax, var_28
  loc_00610644: lea edx, var_2C
  loc_00610647: push edx
  loc_00610648: push eax
  loc_00610649: mov ecx, [eax]
  loc_0061064B: mov esi, eax
  loc_0061064D: call [ecx+0000003Ch]
  loc_00610650: test eax, eax
  loc_00610652: fnclex
  loc_00610654: jge 00610661h
  loc_00610656: push 0000003Ch
  loc_00610658: push 00439BF8h
  loc_0061065D: push esi
  loc_0061065E: push eax
  loc_0061065F: call edi
  loc_00610661: lea ebx, var_30
  loc_00610664: mov eax, var_2C
  loc_00610667: push ebx
  loc_00610668: mov edx, 00000008h
  loc_0061066D: sub esp, 00000010h
  loc_00610670: mov esi, [eax]
  loc_00610672: mov ebx, esp
  loc_00610674: mov ecx, 0043B888h ; "LabelAlamat"
  loc_00610679: push eax
  loc_0061067A: mov var_BC, eax
  loc_00610680: mov [ebx], edx
  loc_00610682: mov edx, var_7C
  loc_00610685: mov [ebx+00000004h], edx
  loc_00610688: mov [ebx+00000008h], ecx
  loc_0061068B: mov ecx, var_74
  loc_0061068E: mov [ebx+0000000Ch], ecx
  loc_00610691: call [esi+0000001Ch]
  loc_00610694: test eax, eax
  loc_00610696: fnclex
  loc_00610698: jge 006106ABh
  loc_0061069A: mov edx, var_BC
  loc_006106A0: push 0000001Ch
  loc_006106A2: push 0043B7D8h
  loc_006106A7: push edx
  loc_006106A8: push eax
  loc_006106A9: call edi
  loc_006106AB: mov ecx, [006680E8h]
  loc_006106B1: mov edx, [006680ECh]
  loc_006106B7: sub esp, 00000010h
  loc_006106BA: mov eax, esp
  loc_006106BC: push 0043B85Ch ; "Caption"
  loc_006106C1: mov [eax], ecx
  loc_006106C3: mov ecx, [006680F0h]
  loc_006106C9: mov [eax+00000004h], edx
  loc_006106CC: mov edx, [006680F4h]
  loc_006106D2: mov [eax+00000008h], ecx
  loc_006106D5: mov [eax+0000000Ch], edx
  loc_006106D8: mov eax, var_30
  loc_006106DB: push eax
  loc_006106DC: call [00401094h] ; __vbaLateMemSt
  loc_006106E2: lea ecx, var_30
  loc_006106E5: lea edx, var_2C
  loc_006106E8: push ecx
  loc_006106E9: lea eax, var_28
  loc_006106EC: push edx
  loc_006106ED: lea ecx, var_24
  loc_006106F0: push eax
  loc_006106F1: push ecx
  loc_006106F2: push 00000004h
  loc_006106F4: call [0040104Ch] ; __vbaFreeObjList
  loc_006106FA: mov eax, var_C4
  loc_00610700: add esp, 00000014h
  loc_00610703: lea ecx, var_24
  loc_00610706: mov edx, [eax]
  loc_00610708: push ecx
  loc_00610709: push eax
  loc_0061070A: call [edx+0000008Ch]
  loc_00610710: test eax, eax
  loc_00610712: fnclex
  loc_00610714: jge 0061072Ah
  loc_00610716: mov edx, var_C4
  loc_0061071C: push 0000008Ch
  loc_00610721: push 0043B5D0h
  loc_00610726: push edx
  loc_00610727: push eax
  loc_00610728: call edi
  loc_0061072A: lea ebx, var_28
  loc_0061072D: mov eax, var_24
  loc_00610730: push ebx
  loc_00610731: mov edx, 00000008h
  loc_00610736: sub esp, 00000010h
  loc_00610739: mov var_70, edx
  loc_0061073C: mov ebx, esp
  loc_0061073E: mov ecx, 0043B848h ; "Section4"
  loc_00610743: mov var_68, ecx
  loc_00610746: mov esi, [eax]
  loc_00610748: mov [ebx], edx
  loc_0061074A: mov edx, var_6C
  loc_0061074D: push eax
  loc_0061074E: mov var_AC, eax
  loc_00610754: mov [ebx+00000004h], edx
  loc_00610757: mov [ebx+00000008h], ecx
  loc_0061075A: mov ecx, var_64
  loc_0061075D: mov [ebx+0000000Ch], ecx
  loc_00610760: call [esi+0000001Ch]
  loc_00610763: test eax, eax
  loc_00610765: fnclex
  loc_00610767: jge 0061077Ah
  loc_00610769: mov edx, var_AC
  loc_0061076F: push 0000001Ch
  loc_00610771: push 00439DD8h
  loc_00610776: push edx
  loc_00610777: push eax
  loc_00610778: call edi
  loc_0061077A: mov eax, var_28
  loc_0061077D: lea edx, var_2C
  loc_00610780: push edx
  loc_00610781: push eax
  loc_00610782: mov ecx, [eax]
  loc_00610784: mov esi, eax
  loc_00610786: call [ecx+0000003Ch]
  loc_00610789: test eax, eax
  loc_0061078B: fnclex
  loc_0061078D: jge 0061079Ah
  loc_0061078F: push 0000003Ch
  loc_00610791: push 00439BF8h
  loc_00610796: push esi
  loc_00610797: push eax
  loc_00610798: call edi
  loc_0061079A: lea ebx, var_30
  loc_0061079D: mov eax, var_2C
  loc_006107A0: push ebx
  loc_006107A1: mov edx, 00000008h
  loc_006107A6: sub esp, 00000010h
  loc_006107A9: mov esi, [eax]
  loc_006107AB: mov ebx, esp
  loc_006107AD: mov ecx, 0043B8A4h ; "LabelKota"
  loc_006107B2: push eax
  loc_006107B3: mov var_BC, eax
  loc_006107B9: mov [ebx], edx
  loc_006107BB: mov edx, var_7C
  loc_006107BE: mov [ebx+00000004h], edx
  loc_006107C1: mov [ebx+00000008h], ecx
  loc_006107C4: mov ecx, var_74
  loc_006107C7: mov [ebx+0000000Ch], ecx
  loc_006107CA: call [esi+0000001Ch]
  loc_006107CD: test eax, eax
  loc_006107CF: fnclex
  loc_006107D1: jge 006107E4h
  loc_006107D3: mov edx, var_BC
  loc_006107D9: push 0000001Ch
  loc_006107DB: push 0043B7D8h
  loc_006107E0: push edx
  loc_006107E1: push eax
  loc_006107E2: call edi
  loc_006107E4: mov ecx, [006680F8h]
  loc_006107EA: mov edx, [006680FCh]
  loc_006107F0: sub esp, 00000010h
  loc_006107F3: mov eax, esp
  loc_006107F5: push 0043B85Ch ; "Caption"
  loc_006107FA: mov [eax], ecx
  loc_006107FC: mov ecx, [00668100h]
  loc_00610802: mov [eax+00000004h], edx
  loc_00610805: mov edx, [00668104h]
  loc_0061080B: mov [eax+00000008h], ecx
  loc_0061080E: mov [eax+0000000Ch], edx
  loc_00610811: mov eax, var_30
  loc_00610814: push eax
  loc_00610815: call [00401094h] ; __vbaLateMemSt
  loc_0061081B: lea ecx, var_30
  loc_0061081E: lea edx, var_2C
  loc_00610821: push ecx
  loc_00610822: lea eax, var_28
  loc_00610825: push edx
  loc_00610826: lea ecx, var_24
  loc_00610829: push eax
  loc_0061082A: push ecx
  loc_0061082B: push 00000004h
  loc_0061082D: call [0040104Ch] ; __vbaFreeObjList
  loc_00610833: mov eax, var_C4
  loc_00610839: add esp, 00000014h
  loc_0061083C: lea ecx, var_24
  loc_0061083F: mov edx, [eax]
  loc_00610841: push ecx
  loc_00610842: push eax
  loc_00610843: call [edx+0000008Ch]
  loc_00610849: test eax, eax
  loc_0061084B: fnclex
  loc_0061084D: jge 00610863h
  loc_0061084F: mov edx, var_C4
  loc_00610855: push 0000008Ch
  loc_0061085A: push 0043B5D0h
  loc_0061085F: push edx
  loc_00610860: push eax
  loc_00610861: call edi
  loc_00610863: lea ebx, var_28
  loc_00610866: mov eax, var_24
  loc_00610869: push ebx
  loc_0061086A: mov edx, 00000008h
  loc_0061086F: sub esp, 00000010h
  loc_00610872: mov var_70, edx
  loc_00610875: mov ebx, esp
  loc_00610877: mov ecx, 0043B848h ; "Section4"
  loc_0061087C: mov var_68, ecx
  loc_0061087F: mov esi, [eax]
  loc_00610881: mov [ebx], edx
  loc_00610883: mov edx, var_6C
  loc_00610886: push eax
  loc_00610887: mov var_AC, eax
  loc_0061088D: mov [ebx+00000004h], edx
  loc_00610890: mov [ebx+00000008h], ecx
  loc_00610893: mov ecx, var_64
  loc_00610896: mov [ebx+0000000Ch], ecx
  loc_00610899: call [esi+0000001Ch]
  loc_0061089C: test eax, eax
  loc_0061089E: fnclex
  loc_006108A0: jge 006108B3h
  loc_006108A2: mov edx, var_AC
  loc_006108A8: push 0000001Ch
  loc_006108AA: push 00439DD8h
  loc_006108AF: push edx
  loc_006108B0: push eax
  loc_006108B1: call edi
  loc_006108B3: mov eax, var_28
  loc_006108B6: lea edx, var_2C
  loc_006108B9: push edx
  loc_006108BA: push eax
  loc_006108BB: mov ecx, [eax]
  loc_006108BD: mov esi, eax
  loc_006108BF: call [ecx+0000003Ch]
  loc_006108C2: test eax, eax
  loc_006108C4: fnclex
  loc_006108C6: jge 006108D3h
  loc_006108C8: push 0000003Ch
  loc_006108CA: push 00439BF8h
  loc_006108CF: push esi
  loc_006108D0: push eax
  loc_006108D1: call edi
  loc_006108D3: lea ebx, var_30
  loc_006108D6: mov eax, var_2C
  loc_006108D9: push ebx
  loc_006108DA: mov edx, 00000008h
  loc_006108DF: sub esp, 00000010h
  loc_006108E2: mov esi, [eax]
  loc_006108E4: mov ebx, esp
  loc_006108E6: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_006108EB: push eax
  loc_006108EC: mov var_BC, eax
  loc_006108F2: mov [ebx], edx
  loc_006108F4: mov edx, var_7C
  loc_006108F7: mov [ebx+00000004h], edx
  loc_006108FA: mov [ebx+00000008h], ecx
  loc_006108FD: mov ecx, var_74
  loc_00610900: mov [ebx+0000000Ch], ecx
  loc_00610903: call [esi+0000001Ch]
  loc_00610906: test eax, eax
  loc_00610908: fnclex
  loc_0061090A: jge 0061091Dh
  loc_0061090C: mov edx, var_BC
  loc_00610912: push 0000001Ch
  loc_00610914: push 0043B7D8h
  loc_00610919: push edx
  loc_0061091A: push eax
  loc_0061091B: call edi
  loc_0061091D: mov ecx, [00668108h]
  loc_00610923: mov edx, [0066810Ch]
  loc_00610929: sub esp, 00000010h
  loc_0061092C: mov eax, esp
  loc_0061092E: push 0043B85Ch ; "Caption"
  loc_00610933: mov [eax], ecx
  loc_00610935: mov ecx, [00668110h]
  loc_0061093B: mov [eax+00000004h], edx
  loc_0061093E: mov edx, [00668114h]
  loc_00610944: mov [eax+00000008h], ecx
  loc_00610947: mov [eax+0000000Ch], edx
  loc_0061094A: mov eax, var_30
  loc_0061094D: push eax
  loc_0061094E: call [00401094h] ; __vbaLateMemSt
  loc_00610954: lea ecx, var_30
  loc_00610957: lea edx, var_2C
  loc_0061095A: push ecx
  loc_0061095B: lea eax, var_28
  loc_0061095E: push edx
  loc_0061095F: lea ecx, var_24
  loc_00610962: push eax
  loc_00610963: push ecx
  loc_00610964: push 00000004h
  loc_00610966: call [0040104Ch] ; __vbaFreeObjList
  loc_0061096C: mov eax, var_C4
  loc_00610972: add esp, 00000014h
  loc_00610975: mov edx, [eax]
  loc_00610977: push 0000000Ah
  loc_00610979: push eax
  loc_0061097A: call [edx+00000048h]
  loc_0061097D: test eax, eax
  loc_0061097F: fnclex
  loc_00610981: jge 00610994h
  loc_00610983: mov ecx, var_C4
  loc_00610989: push 00000048h
  loc_0061098B: push 0043B5D0h
  loc_00610990: push ecx
  loc_00610991: push eax
  loc_00610992: call edi
  loc_00610994: mov eax, var_C4
  loc_0061099A: push 0000000Ah
  loc_0061099C: push eax
  loc_0061099D: mov edx, [eax]
  loc_0061099F: call [edx+00000058h]
  loc_006109A2: test eax, eax
  loc_006109A4: fnclex
  loc_006109A6: jge 006109B9h
  loc_006109A8: mov ecx, var_C4
  loc_006109AE: push 00000058h
  loc_006109B0: push 0043B5D0h
  loc_006109B5: push ecx
  loc_006109B6: push eax
  loc_006109B7: call edi
  loc_006109B9: sub esp, 00000010h
  loc_006109BC: mov eax, 00000002h
  loc_006109C1: mov edx, esp
  loc_006109C3: mov ecx, eax
  loc_006109C5: mov var_70, ecx
  loc_006109C8: mov var_68, eax
  loc_006109CB: mov [edx], ecx
  loc_006109CD: mov ecx, var_6C
  loc_006109D0: push 8001000Eh
  loc_006109D5: mov [edx+00000004h], ecx
  loc_006109D8: mov ecx, var_C4
  loc_006109DE: push ecx
  loc_006109DF: mov [edx+00000008h], eax
  loc_006109E2: mov eax, var_64
  loc_006109E5: mov [edx+0000000Ch], eax
  loc_006109E8: call [00401280h] ; __vbaLateIdSt
  loc_006109EE: mov edx, var_C4
  loc_006109F4: push 00000000h
  loc_006109F6: push 80011003h
  loc_006109FB: push edx
  loc_006109FC: call [0040102Ch] ; __vbaLateIdCall
  loc_00610A02: add esp, 0000000Ch
  loc_00610A05: lea eax, var_C4
  loc_00610A0B: push 00000000h
  loc_00610A0D: push eax
  loc_00610A0E: call [004010C8h] ; __vbaObjSetAddref
  loc_00610A14: push 00610A7Eh
  loc_00610A19: jmp 00610A54h
  loc_00610A1B: lea ecx, var_20
  loc_00610A1E: call [0040129Ch] ; __vbaFreeStr
  loc_00610A24: lea ecx, var_30
  loc_00610A27: lea edx, var_2C
  loc_00610A2A: push ecx
  loc_00610A2B: lea eax, var_28
  loc_00610A2E: push edx
  loc_00610A2F: lea ecx, var_24
  loc_00610A32: push eax
  loc_00610A33: push ecx
  loc_00610A34: push 00000004h
  loc_00610A36: call [0040104Ch] ; __vbaFreeObjList
  loc_00610A3C: lea edx, var_60
  loc_00610A3F: lea eax, var_50
  loc_00610A42: push edx
  loc_00610A43: lea ecx, var_40
  loc_00610A46: push eax
  loc_00610A47: push ecx
  loc_00610A48: push 00000003h
  loc_00610A4A: call [0040103Ch] ; __vbaFreeVarList
  loc_00610A50: add esp, 00000024h
  loc_00610A53: ret
  loc_00610A54: lea edx, var_C8
  loc_00610A5A: lea eax, var_C4
  loc_00610A60: push edx
  loc_00610A61: push eax
  loc_00610A62: push 00000002h
  loc_00610A64: call [0040104Ch] ; __vbaFreeObjList
  loc_00610A6A: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_00610A70: add esp, 0000000Ch
  loc_00610A73: lea ecx, var_18
  loc_00610A76: call __vbaFreeStr
  loc_00610A78: lea ecx, var_1C
  loc_00610A7B: call __vbaFreeStr
  loc_00610A7D: ret
  loc_00610A7E: mov ecx, var_10
  loc_00610A81: pop edi
  loc_00610A82: pop esi
  loc_00610A83: mov fs:[00000000h], ecx
  loc_00610A8A: pop ebx
  loc_00610A8B: mov esp, ebp
  loc_00610A8D: pop ebp
  loc_00610A8E: ret
End Sub

Private Sub Proc_34_8_610AA0
  loc_00610AA0: push ebp
  loc_00610AA1: mov ebp, esp
  loc_00610AA3: sub esp, 00000008h
  loc_00610AA6: push 00405696h ; __vbaExceptHandler
  loc_00610AAB: mov eax, fs:[00000000h]
  loc_00610AB1: push eax
  loc_00610AB2: mov fs:[00000000h], esp
  loc_00610AB9: sub esp, 000000D0h
  loc_00610ABF: push ebx
  loc_00610AC0: push esi
  loc_00610AC1: push edi
  loc_00610AC2: mov var_8, esp
  loc_00610AC5: mov var_4, 00404520h
  loc_00610ACC: xor ebx, ebx
  loc_00610ACE: mov var_18, ebx
  loc_00610AD1: mov var_1C, ebx
  loc_00610AD4: mov var_20, ebx
  loc_00610AD7: mov var_24, ebx
  loc_00610ADA: mov var_28, ebx
  loc_00610ADD: mov var_2C, ebx
  loc_00610AE0: mov var_30, ebx
  loc_00610AE3: mov var_40, ebx
  loc_00610AE6: mov var_50, ebx
  loc_00610AE9: mov var_60, ebx
  loc_00610AEC: mov var_70, ebx
  loc_00610AEF: mov var_80, ebx
  loc_00610AF2: mov var_90, ebx
  loc_00610AF8: mov var_A4, ebx
  loc_00610AFE: mov var_C4, ebx
  loc_00610B04: mov var_C8, ebx
  loc_00610B0A: call 0055B320h
  loc_00610B0F: mov edx, 0042DFECh
  loc_00610B14: lea ecx, var_18
  loc_00610B17: call [004011FCh] ; __vbaStrCopy
  loc_00610B1D: push 0043C030h ; " SELECT No_Retur_beli, no_pembelian, no_notabeli, tgl_retur_beli, nama_pemasok,"
  loc_00610B22: push 0043C0E4h ; "tot_retur_beli FROM Retur_beli order BY No_Retur_beli"
  loc_00610B27: call [00401070h] ; __vbaStrCat
  loc_00610B2D: mov edx, eax
  loc_00610B2F: lea ecx, var_18
  loc_00610B32: call [0040126Ch] ; __vbaStrMove
  loc_00610B38: push 0042DF30h
  loc_00610B3D: call [00401168h] ; __vbaNew
  loc_00610B43: push eax
  loc_00610B44: push 00668090h
  loc_00610B49: call [004010B8h] ; __vbaObjSet
  loc_00610B4F: cmp [0066803Ch], ebx
  loc_00610B55: jnz 00610B67h
  loc_00610B57: push 0066803Ch
  loc_00610B5C: push 0042DEFCh
  loc_00610B61: call [004011E8h] ; __vbaNew2
  loc_00610B67: mov esi, [0066803Ch]
  loc_00610B6D: lea ecx, var_40
  loc_00610B70: mov var_38, 80020004h
  loc_00610B77: mov var_40, 0000000Ah
  loc_00610B7E: call [0040123Ch] ; __vbaFreeVarg
  loc_00610B84: mov eax, [esi]
  loc_00610B86: lea ecx, var_24
  loc_00610B89: push ecx
  loc_00610B8A: mov ecx, var_18
  loc_00610B8D: lea edx, var_40
  loc_00610B90: push 00000001h
  loc_00610B92: push edx
  loc_00610B93: push ecx
  loc_00610B94: push esi
  loc_00610B95: call [eax+00000040h]
  loc_00610B98: cmp eax, ebx
  loc_00610B9A: fnclex
  loc_00610B9C: jge 00610BB1h
  loc_00610B9E: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00610BA4: push 00000040h
  loc_00610BA6: push 0042E1B0h
  loc_00610BAB: push esi
  loc_00610BAC: push eax
  loc_00610BAD: call edi
  loc_00610BAF: jmp 00610BB7h
  loc_00610BB1: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00610BB7: mov edx, var_24
  loc_00610BBA: mov esi, [00401268h] ; __vbaCastObj
  loc_00610BC0: push 0042DF20h
  loc_00610BC5: push edx
  loc_00610BC6: call __vbaCastObj
  loc_00610BC8: push eax
  loc_00610BC9: push 00668090h
  loc_00610BCE: call [004010B8h] ; __vbaObjSet
  loc_00610BD4: lea ecx, var_24
  loc_00610BD7: call [004012A0h] ; __vbaFreeObj
  loc_00610BDD: lea ecx, var_40
  loc_00610BE0: call [00401020h] ; __vbaFreeVar
  loc_00610BE6: cmp [00668594h], ebx
  loc_00610BEC: jnz 00610BFEh
  loc_00610BEE: push 00668594h
  loc_00610BF3: push 004078F8h
  loc_00610BF8: call [004011E8h] ; __vbaNew2
  loc_00610BFE: mov eax, [00668594h]
  loc_00610C03: lea ecx, var_C4
  loc_00610C09: push eax
  loc_00610C0A: push ecx
  loc_00610C0B: call [004010C8h] ; __vbaObjSetAddref
  loc_00610C11: mov edx, var_C4
  loc_00610C17: push 0043B7B8h
  loc_00610C1C: push ebx
  loc_00610C1D: mov edx, [edx]
  loc_00610C1F: mov var_E0, edx
  loc_00610C25: call __vbaCastObj
  loc_00610C27: push eax
  loc_00610C28: lea eax, var_24
  loc_00610C2B: push eax
  loc_00610C2C: call [004010B8h] ; __vbaObjSet
  loc_00610C32: mov ecx, var_C4
  loc_00610C38: mov edx, var_E0
  loc_00610C3E: push eax
  loc_00610C3F: push ecx
  loc_00610C40: call [edx+00000028h]
  loc_00610C43: cmp eax, ebx
  loc_00610C45: fnclex
  loc_00610C47: jge 00610C5Ah
  loc_00610C49: mov ecx, var_C4
  loc_00610C4F: push 00000028h
  loc_00610C51: push 0043B5D0h
  loc_00610C56: push ecx
  loc_00610C57: push eax
  loc_00610C58: call edi
  loc_00610C5A: lea ecx, var_24
  loc_00610C5D: call [004012A0h] ; __vbaFreeObj
  loc_00610C63: mov eax, var_C4
  loc_00610C69: push 0042DFECh
  loc_00610C6E: push eax
  loc_00610C6F: mov edx, [eax]
  loc_00610C71: call [edx+00000030h]
  loc_00610C74: cmp eax, ebx
  loc_00610C76: fnclex
  loc_00610C78: jge 00610C8Bh
  loc_00610C7A: mov ecx, var_C4
  loc_00610C80: push 00000030h
  loc_00610C82: push 0043B5D0h
  loc_00610C87: push ecx
  loc_00610C88: push eax
  loc_00610C89: call edi
  loc_00610C8B: mov eax, [00668090h]
  loc_00610C90: lea ecx, var_24
  loc_00610C93: push ecx
  loc_00610C94: push eax
  loc_00610C95: mov edx, [eax]
  loc_00610C97: call [edx+00000114h]
  loc_00610C9D: cmp eax, ebx
  loc_00610C9F: fnclex
  loc_00610CA1: jge 00610CB7h
  loc_00610CA3: mov edx, [00668090h]
  loc_00610CA9: push 00000114h
  loc_00610CAE: push 0043B7C8h
  loc_00610CB3: push edx
  loc_00610CB4: push eax
  loc_00610CB5: call edi
  loc_00610CB7: mov eax, var_C4
  loc_00610CBD: mov ecx, var_24
  loc_00610CC0: push 0043B7B8h
  loc_00610CC5: push ecx
  loc_00610CC6: mov esi, [eax]
  loc_00610CC8: call [00401268h] ; __vbaCastObj
  loc_00610CCE: lea edx, var_28
  loc_00610CD1: push eax
  loc_00610CD2: push edx
  loc_00610CD3: call [004010B8h] ; __vbaObjSet
  loc_00610CD9: push eax
  loc_00610CDA: mov eax, var_C4
  loc_00610CE0: push eax
  loc_00610CE1: call [esi+00000028h]
  loc_00610CE4: cmp eax, ebx
  loc_00610CE6: fnclex
  loc_00610CE8: jge 00610CFBh
  loc_00610CEA: mov ecx, var_C4
  loc_00610CF0: push 00000028h
  loc_00610CF2: push 0043B5D0h
  loc_00610CF7: push ecx
  loc_00610CF8: push eax
  loc_00610CF9: call edi
  loc_00610CFB: lea edx, var_28
  loc_00610CFE: lea eax, var_24
  loc_00610D01: push edx
  loc_00610D02: push eax
  loc_00610D03: push 00000002h
  loc_00610D05: call [0040104Ch] ; __vbaFreeObjList
  loc_00610D0B: mov eax, var_C4
  loc_00610D11: add esp, 0000000Ch
  loc_00610D14: lea edx, var_24
  loc_00610D17: mov ecx, [eax]
  loc_00610D19: push edx
  loc_00610D1A: push eax
  loc_00610D1B: call [ecx+0000008Ch]
  loc_00610D21: cmp eax, ebx
  loc_00610D23: fnclex
  loc_00610D25: jge 00610D3Bh
  loc_00610D27: mov ecx, var_C4
  loc_00610D2D: push 0000008Ch
  loc_00610D32: push 0043B5D0h
  loc_00610D37: push ecx
  loc_00610D38: push eax
  loc_00610D39: call edi
  loc_00610D3B: lea esi, var_28
  loc_00610D3E: mov eax, var_24
  loc_00610D41: push esi
  loc_00610D42: mov ecx, 00000008h
  loc_00610D47: sub esp, 00000010h
  loc_00610D4A: mov var_70, ecx
  loc_00610D4D: mov esi, esp
  loc_00610D4F: mov var_68, 00439FB0h ; "Section1"
  loc_00610D56: mov edx, [eax]
  loc_00610D58: push eax
  loc_00610D59: mov [esi], ecx
  loc_00610D5B: mov ecx, var_6C
  loc_00610D5E: mov var_AC, eax
  loc_00610D64: mov [esi+00000004h], ecx
  loc_00610D67: mov ecx, var_68
  loc_00610D6A: mov [esi+00000008h], ecx
  loc_00610D6D: mov ecx, var_64
  loc_00610D70: mov [esi+0000000Ch], ecx
  loc_00610D73: call [edx+0000001Ch]
  loc_00610D76: cmp eax, ebx
  loc_00610D78: fnclex
  loc_00610D7A: jge 00610D8Dh
  loc_00610D7C: mov edx, var_AC
  loc_00610D82: push 0000001Ch
  loc_00610D84: push 00439DD8h
  loc_00610D89: push edx
  loc_00610D8A: push eax
  loc_00610D8B: call edi
  loc_00610D8D: mov eax, var_28
  loc_00610D90: lea edx, var_2C
  loc_00610D93: push edx
  loc_00610D94: push eax
  loc_00610D95: mov ecx, [eax]
  loc_00610D97: mov esi, eax
  loc_00610D99: call [ecx+0000003Ch]
  loc_00610D9C: cmp eax, ebx
  loc_00610D9E: fnclex
  loc_00610DA0: jge 00610DADh
  loc_00610DA2: push 0000003Ch
  loc_00610DA4: push 00439BF8h
  loc_00610DA9: push esi
  loc_00610DAA: push eax
  loc_00610DAB: call edi
  loc_00610DAD: mov eax, var_2C
  loc_00610DB0: mov var_2C, ebx
  loc_00610DB3: push eax
  loc_00610DB4: lea eax, var_C8
  loc_00610DBA: push eax
  loc_00610DBB: call [004010B8h] ; __vbaObjSet
  loc_00610DC1: lea ecx, var_28
  loc_00610DC4: lea edx, var_24
  loc_00610DC7: push ecx
  loc_00610DC8: push edx
  loc_00610DC9: push 00000002h
  loc_00610DCB: call [0040104Ch] ; __vbaFreeObjList
  loc_00610DD1: mov eax, var_C8
  loc_00610DD7: add esp, 0000000Ch
  loc_00610DDA: lea edx, var_A4
  loc_00610DE0: mov ecx, [eax]
  loc_00610DE2: push edx
  loc_00610DE3: push eax
  loc_00610DE4: call [ecx+00000020h]
  loc_00610DE7: cmp eax, ebx
  loc_00610DE9: fnclex
  loc_00610DEB: jge 00610DFEh
  loc_00610DED: mov ecx, var_C8
  loc_00610DF3: push 00000020h
  loc_00610DF5: push 0043B7D8h
  loc_00610DFA: push ecx
  loc_00610DFB: push eax
  loc_00610DFC: call edi
  loc_00610DFE: mov ecx, var_A4
  loc_00610E04: call [00401138h] ; __vbaI2I4
  loc_00610E0A: mov var_D0, eax
  loc_00610E10: mov ebx, 00000001h
  loc_00610E15: cmp bx, var_D0
  loc_00610E1C: mov var_14, ebx
  loc_00610E1F: jg 00611090h
  loc_00610E25: lea esi, var_24
  loc_00610E28: mov ecx, var_C8
  loc_00610E2E: push esi
  loc_00610E2F: mov eax, 00000002h
  loc_00610E34: sub esp, 00000010h
  loc_00610E37: mov var_70, eax
  loc_00610E3A: mov esi, esp
  loc_00610E3C: mov var_68, bx
  loc_00610E40: mov edx, [ecx]
  loc_00610E42: push ecx
  loc_00610E43: mov [esi], eax
  loc_00610E45: mov eax, var_6C
  loc_00610E48: mov [esi+00000004h], eax
  loc_00610E4B: mov eax, var_68
  loc_00610E4E: mov [esi+00000008h], eax
  loc_00610E51: mov eax, var_64
  loc_00610E54: mov [esi+0000000Ch], eax
  loc_00610E57: call [edx+0000001Ch]
  loc_00610E5A: test eax, eax
  loc_00610E5C: fnclex
  loc_00610E5E: jge 00610E71h
  loc_00610E60: mov ecx, var_C8
  loc_00610E66: push 0000001Ch
  loc_00610E68: push 0043B7D8h
  loc_00610E6D: push ecx
  loc_00610E6E: push eax
  loc_00610E6F: call edi
  loc_00610E71: mov edx, var_24
  loc_00610E74: push 0043B7E8h
  loc_00610E79: push edx
  loc_00610E7A: call [004011CCh] ; __vbaCheckType
  loc_00610E80: lea ecx, var_24
  loc_00610E83: mov si, ax
  loc_00610E86: call [004012A0h] ; __vbaFreeObj
  loc_00610E8C: test si, si
  loc_00610E8F: Unknown_795()
  loc_00610E95: mov var_68, bx
  loc_00610E99: lea ebx, var_24
  loc_00610E9C: push ebx
  loc_00610E9D: mov ecx, var_C8
  loc_00610EA3: sub esp, 00000010h
  loc_00610EA6: mov eax, 00000002h
  loc_00610EAB: mov ebx, esp
  loc_00610EAD: mov var_70, eax
  loc_00610EB0: mov edx, [ecx]
  loc_00610EB2: mov esi, 0042DFECh
  loc_00610EB7: mov [ebx], eax
  loc_00610EB9: mov eax, var_6C
  loc_00610EBC: push ecx
  loc_00610EBD: mov var_78, esi
  loc_00610EC0: mov [ebx+00000004h], eax
  loc_00610EC3: mov eax, var_68
  loc_00610EC6: mov edi, 00000008h
  loc_00610ECB: mov [ebx+00000008h], eax
  loc_00610ECE: mov eax, var_64
  loc_00610ED1: mov [ebx+0000000Ch], eax
  loc_00610ED4: call [edx+0000001Ch]
  loc_00610ED7: test eax, eax
  loc_00610ED9: fnclex
  loc_00610EDB: jge 00610EF2h
  loc_00610EDD: mov ecx, var_C8
  loc_00610EE3: push 0000001Ch
  loc_00610EE5: push 0043B7D8h
  loc_00610EEA: push ecx
  loc_00610EEB: push eax
  loc_00610EEC: call [00401080h] ; __vbaHresultCheckObj
  loc_00610EF2: mov eax, var_7C
  loc_00610EF5: sub esp, 00000010h
  loc_00610EF8: mov edx, esp
  loc_00610EFA: mov ecx, var_74
  loc_00610EFD: push 0043B7F8h ; "DataMember"
  loc_00610F02: mov [edx], edi
  loc_00610F04: mov [edx+00000004h], eax
  loc_00610F07: mov [edx+00000008h], esi
  loc_00610F0A: mov [edx+0000000Ch], ecx
  loc_00610F0D: mov edx, var_24
  loc_00610F10: push edx
  loc_00610F11: call [00401094h] ; __vbaLateMemSt
  loc_00610F17: lea ecx, var_24
  loc_00610F1A: call [004012A0h] ; __vbaFreeObj
  loc_00610F20: mov eax, [00668090h]
  loc_00610F25: lea edx, var_24
  loc_00610F28: push edx
  loc_00610F29: push eax
  loc_00610F2A: mov ecx, [eax]
  loc_00610F2C: call [ecx+00000054h]
  loc_00610F2F: test eax, eax
  loc_00610F31: fnclex
  loc_00610F33: jge 00610F4Ah
  loc_00610F35: mov ecx, [00668090h]
  loc_00610F3B: push 00000054h
  loc_00610F3D: push 0042DF44h
  loc_00610F42: push ecx
  loc_00610F43: push eax
  loc_00610F44: call [00401080h] ; __vbaHresultCheckObj
  loc_00610F4A: mov ebx, var_14
  loc_00610F4D: lea edi, var_28
  loc_00610F50: mov dx, bx
  loc_00610F53: push edi
  loc_00610F54: sub dx, 0001h
  loc_00610F58: mov eax, var_24
  loc_00610F5B: jo 0061181Eh
  loc_00610F61: sub esp, 00000010h
  loc_00610F64: mov ecx, 00000002h
  loc_00610F69: mov edi, esp
  loc_00610F6B: mov var_70, ecx
  loc_00610F6E: mov var_68, dx
  loc_00610F72: mov edx, [eax]
  loc_00610F74: mov [edi], ecx
  loc_00610F76: mov ecx, var_6C
  loc_00610F79: push eax
  loc_00610F7A: mov esi, eax
  loc_00610F7C: mov [edi+00000004h], ecx
  loc_00610F7F: mov ecx, var_68
  loc_00610F82: mov [edi+00000008h], ecx
  loc_00610F85: mov ecx, var_64
  loc_00610F88: mov [edi+0000000Ch], ecx
  loc_00610F8B: call [edx+00000028h]
  loc_00610F8E: test eax, eax
  loc_00610F90: fnclex
  loc_00610F92: jge 00610FA7h
  loc_00610F94: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00610F9A: push 00000028h
  loc_00610F9C: push 0042DFACh
  loc_00610FA1: push esi
  loc_00610FA2: push eax
  loc_00610FA3: call edi
  loc_00610FA5: jmp 00610FADh
  loc_00610FA7: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00610FAD: mov eax, var_28
  loc_00610FB0: lea ecx, var_20
  loc_00610FB3: push ecx
  loc_00610FB4: push eax
  loc_00610FB5: mov edx, [eax]
  loc_00610FB7: mov esi, eax
  loc_00610FB9: call [edx+0000002Ch]
  loc_00610FBC: test eax, eax
  loc_00610FBE: fnclex
  loc_00610FC0: jge 00610FCDh
  loc_00610FC2: push 0000002Ch
  loc_00610FC4: push 0042DFBCh
  loc_00610FC9: push esi
  loc_00610FCA: push eax
  loc_00610FCB: call edi
  loc_00610FCD: mov eax, var_20
  loc_00610FD0: lea esi, var_2C
  loc_00610FD3: push esi
  loc_00610FD4: mov ecx, var_C8
  loc_00610FDA: sub esp, 00000010h
  loc_00610FDD: mov var_38, eax
  loc_00610FE0: mov esi, esp
  loc_00610FE2: mov eax, 00000002h
  loc_00610FE7: mov var_78, bx
  loc_00610FEB: mov var_20, 00000000h
  loc_00610FF2: mov [esi], eax
  loc_00610FF4: mov eax, var_7C
  loc_00610FF7: mov var_40, 00000008h
  loc_00610FFE: mov edx, [ecx]
  loc_00611000: mov [esi+00000004h], eax
  loc_00611003: mov eax, var_78
  loc_00611006: push ecx
  loc_00611007: mov [esi+00000008h], eax
  loc_0061100A: mov eax, var_74
  loc_0061100D: mov [esi+0000000Ch], eax
  loc_00611010: call [edx+0000001Ch]
  loc_00611013: test eax, eax
  loc_00611015: fnclex
  loc_00611017: jge 0061102Ah
  loc_00611019: mov ecx, var_C8
  loc_0061101F: push 0000001Ch
  loc_00611021: push 0043B7D8h
  loc_00611026: push ecx
  loc_00611027: push eax
  loc_00611028: call edi
  loc_0061102A: mov eax, var_40
  loc_0061102D: mov ecx, var_3C
  loc_00611030: sub esp, 00000010h
  loc_00611033: mov edx, esp
  loc_00611035: push 0043B810h ; "DataField"
  loc_0061103A: mov [edx], eax
  loc_0061103C: mov eax, var_38
  loc_0061103F: mov [edx+00000004h], ecx
  loc_00611042: mov ecx, var_34
  loc_00611045: mov [edx+00000008h], eax
  loc_00611048: mov [edx+0000000Ch], ecx
  loc_0061104B: mov edx, var_2C
  loc_0061104E: push edx
  loc_0061104F: call [00401094h] ; __vbaLateMemSt
  loc_00611055: lea eax, var_2C
  loc_00611058: lea ecx, var_28
  loc_0061105B: push eax
  loc_0061105C: lea edx, var_24
  loc_0061105F: push ecx
  loc_00611060: push edx
  loc_00611061: push 00000003h
  loc_00611063: call [0040104Ch] ; __vbaFreeObjList
  loc_00611069: add esp, 00000010h
  loc_0061106C: lea ecx, var_40
  loc_0061106F: call [00401020h] ; __vbaFreeVar
  loc_00611075: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_0061107B: mov eax, 00000001h
  loc_00611080: add ax, bx
  loc_00611083: jo 0061181Eh
  loc_00611089: mov ebx, eax
  loc_0061108B: jmp 00610E15h
  loc_00611090: lea eax, var_C8
  loc_00611096: push 00000000h
  loc_00611098: push eax
  loc_00611099: call [004010C8h] ; __vbaObjSetAddref
  loc_0061109F: lea ecx, var_40
  loc_006110A2: push ecx
  loc_006110A3: call [00401220h] ; rtcGetDateVar
  loc_006110A9: lea edx, var_70
  loc_006110AC: lea ecx, var_50
  loc_006110AF: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_006110B6: mov var_70, 00000008h
  loc_006110BD: call [00401238h] ; __vbaVarDup
  loc_006110C3: push 00000001h
  loc_006110C5: lea edx, var_50
  loc_006110C8: push 00000001h
  loc_006110CA: lea eax, var_40
  loc_006110CD: push edx
  loc_006110CE: lea ecx, var_60
  loc_006110D1: push eax
  loc_006110D2: push ecx
  loc_006110D3: call [00401078h] ; rtcVarFromFormatVar
  loc_006110D9: mov eax, var_C4
  loc_006110DF: lea ecx, var_24
  loc_006110E2: push ecx
  loc_006110E3: push eax
  loc_006110E4: mov edx, [eax]
  loc_006110E6: call [edx+0000008Ch]
  loc_006110EC: test eax, eax
  loc_006110EE: fnclex
  loc_006110F0: jge 00611106h
  loc_006110F2: mov edx, var_C4
  loc_006110F8: push 0000008Ch
  loc_006110FD: push 0043B5D0h
  loc_00611102: push edx
  loc_00611103: push eax
  loc_00611104: call edi
  loc_00611106: lea ebx, var_28
  loc_00611109: mov eax, var_24
  loc_0061110C: push ebx
  loc_0061110D: mov edx, 00000008h
  loc_00611112: sub esp, 00000010h
  loc_00611115: mov esi, [eax]
  loc_00611117: mov ebx, esp
  loc_00611119: mov ecx, 0043B848h ; "Section4"
  loc_0061111E: push eax
  loc_0061111F: mov var_AC, eax
  loc_00611125: mov [ebx], edx
  loc_00611127: mov edx, var_7C
  loc_0061112A: mov [ebx+00000004h], edx
  loc_0061112D: mov [ebx+00000008h], ecx
  loc_00611130: mov ecx, var_74
  loc_00611133: mov [ebx+0000000Ch], ecx
  loc_00611136: call [esi+0000001Ch]
  loc_00611139: test eax, eax
  loc_0061113B: fnclex
  loc_0061113D: jge 00611150h
  loc_0061113F: mov edx, var_AC
  loc_00611145: push 0000001Ch
  loc_00611147: push 00439DD8h
  loc_0061114C: push edx
  loc_0061114D: push eax
  loc_0061114E: call edi
  loc_00611150: mov eax, var_28
  loc_00611153: lea edx, var_2C
  loc_00611156: push edx
  loc_00611157: push eax
  loc_00611158: mov ecx, [eax]
  loc_0061115A: mov esi, eax
  loc_0061115C: call [ecx+0000003Ch]
  loc_0061115F: test eax, eax
  loc_00611161: fnclex
  loc_00611163: jge 00611170h
  loc_00611165: push 0000003Ch
  loc_00611167: push 00439BF8h
  loc_0061116C: push esi
  loc_0061116D: push eax
  loc_0061116E: call edi
  loc_00611170: lea ebx, var_30
  loc_00611173: mov eax, var_2C
  loc_00611176: push ebx
  loc_00611177: mov edx, 00000008h
  loc_0061117C: sub esp, 00000010h
  loc_0061117F: mov esi, [eax]
  loc_00611181: mov ebx, esp
  loc_00611183: mov ecx, 0043B828h ; "labelTanggal"
  loc_00611188: push eax
  loc_00611189: mov var_BC, eax
  loc_0061118F: mov [ebx], edx
  loc_00611191: mov edx, var_8C
  loc_00611197: mov [ebx+00000004h], edx
  loc_0061119A: mov [ebx+00000008h], ecx
  loc_0061119D: mov ecx, var_84
  loc_006111A3: mov [ebx+0000000Ch], ecx
  loc_006111A6: call [esi+0000001Ch]
  loc_006111A9: test eax, eax
  loc_006111AB: fnclex
  loc_006111AD: jge 006111C0h
  loc_006111AF: mov edx, var_BC
  loc_006111B5: push 0000001Ch
  loc_006111B7: push 0043B7D8h
  loc_006111BC: push edx
  loc_006111BD: push eax
  loc_006111BE: call edi
  loc_006111C0: mov ecx, var_60
  loc_006111C3: mov edx, var_5C
  loc_006111C6: sub esp, 00000010h
  loc_006111C9: mov eax, esp
  loc_006111CB: push 0043B85Ch ; "Caption"
  loc_006111D0: mov [eax], ecx
  loc_006111D2: mov ecx, var_58
  loc_006111D5: mov [eax+00000004h], edx
  loc_006111D8: mov edx, var_54
  loc_006111DB: mov [eax+00000008h], ecx
  loc_006111DE: mov [eax+0000000Ch], edx
  loc_006111E1: mov eax, var_30
  loc_006111E4: push eax
  loc_006111E5: call [00401094h] ; __vbaLateMemSt
  loc_006111EB: lea ecx, var_30
  loc_006111EE: lea edx, var_2C
  loc_006111F1: push ecx
  loc_006111F2: lea eax, var_28
  loc_006111F5: push edx
  loc_006111F6: lea ecx, var_24
  loc_006111F9: push eax
  loc_006111FA: push ecx
  loc_006111FB: push 00000004h
  loc_006111FD: call [0040104Ch] ; __vbaFreeObjList
  loc_00611203: lea edx, var_60
  loc_00611206: lea eax, var_50
  loc_00611209: push edx
  loc_0061120A: lea ecx, var_40
  loc_0061120D: push eax
  loc_0061120E: push ecx
  loc_0061120F: push 00000003h
  loc_00611211: call [0040103Ch] ; __vbaFreeVarList
  loc_00611217: mov eax, var_C4
  loc_0061121D: add esp, 00000024h
  loc_00611220: lea ecx, var_24
  loc_00611223: mov edx, [eax]
  loc_00611225: push ecx
  loc_00611226: push eax
  loc_00611227: call [edx+0000008Ch]
  loc_0061122D: test eax, eax
  loc_0061122F: fnclex
  loc_00611231: jge 00611247h
  loc_00611233: mov edx, var_C4
  loc_00611239: push 0000008Ch
  loc_0061123E: push 0043B5D0h
  loc_00611243: push edx
  loc_00611244: push eax
  loc_00611245: call edi
  loc_00611247: lea ebx, var_28
  loc_0061124A: mov eax, var_24
  loc_0061124D: push ebx
  loc_0061124E: mov edx, 00000008h
  loc_00611253: sub esp, 00000010h
  loc_00611256: mov var_70, edx
  loc_00611259: mov ebx, esp
  loc_0061125B: mov ecx, 0043B848h ; "Section4"
  loc_00611260: mov var_68, ecx
  loc_00611263: mov esi, [eax]
  loc_00611265: mov [ebx], edx
  loc_00611267: mov edx, var_6C
  loc_0061126A: push eax
  loc_0061126B: mov var_AC, eax
  loc_00611271: mov [ebx+00000004h], edx
  loc_00611274: mov [ebx+00000008h], ecx
  loc_00611277: mov ecx, var_64
  loc_0061127A: mov [ebx+0000000Ch], ecx
  loc_0061127D: call [esi+0000001Ch]
  loc_00611280: test eax, eax
  loc_00611282: fnclex
  loc_00611284: jge 00611297h
  loc_00611286: mov edx, var_AC
  loc_0061128C: push 0000001Ch
  loc_0061128E: push 00439DD8h
  loc_00611293: push edx
  loc_00611294: push eax
  loc_00611295: call edi
  loc_00611297: mov eax, var_28
  loc_0061129A: lea edx, var_2C
  loc_0061129D: push edx
  loc_0061129E: push eax
  loc_0061129F: mov ecx, [eax]
  loc_006112A1: mov esi, eax
  loc_006112A3: call [ecx+0000003Ch]
  loc_006112A6: test eax, eax
  loc_006112A8: fnclex
  loc_006112AA: jge 006112B7h
  loc_006112AC: push 0000003Ch
  loc_006112AE: push 00439BF8h
  loc_006112B3: push esi
  loc_006112B4: push eax
  loc_006112B5: call edi
  loc_006112B7: lea ebx, var_30
  loc_006112BA: mov eax, var_2C
  loc_006112BD: push ebx
  loc_006112BE: mov edx, 00000008h
  loc_006112C3: sub esp, 00000010h
  loc_006112C6: mov esi, [eax]
  loc_006112C8: mov ebx, esp
  loc_006112CA: mov ecx, 0043B870h ; "LabelNama"
  loc_006112CF: push eax
  loc_006112D0: mov var_BC, eax
  loc_006112D6: mov [ebx], edx
  loc_006112D8: mov edx, var_7C
  loc_006112DB: mov [ebx+00000004h], edx
  loc_006112DE: mov [ebx+00000008h], ecx
  loc_006112E1: mov ecx, var_74
  loc_006112E4: mov [ebx+0000000Ch], ecx
  loc_006112E7: call [esi+0000001Ch]
  loc_006112EA: test eax, eax
  loc_006112EC: fnclex
  loc_006112EE: jge 00611301h
  loc_006112F0: mov edx, var_BC
  loc_006112F6: push 0000001Ch
  loc_006112F8: push 0043B7D8h
  loc_006112FD: push edx
  loc_006112FE: push eax
  loc_006112FF: call edi
  loc_00611301: mov ecx, [006680D8h]
  loc_00611307: mov edx, [006680DCh]
  loc_0061130D: sub esp, 00000010h
  loc_00611310: mov eax, esp
  loc_00611312: push 0043B85Ch ; "Caption"
  loc_00611317: mov [eax], ecx
  loc_00611319: mov ecx, [006680E0h]
  loc_0061131F: mov [eax+00000004h], edx
  loc_00611322: mov edx, [006680E4h]
  loc_00611328: mov [eax+00000008h], ecx
  loc_0061132B: mov [eax+0000000Ch], edx
  loc_0061132E: mov eax, var_30
  loc_00611331: push eax
  loc_00611332: call [00401094h] ; __vbaLateMemSt
  loc_00611338: lea ecx, var_30
  loc_0061133B: lea edx, var_2C
  loc_0061133E: push ecx
  loc_0061133F: lea eax, var_28
  loc_00611342: push edx
  loc_00611343: lea ecx, var_24
  loc_00611346: push eax
  loc_00611347: push ecx
  loc_00611348: push 00000004h
  loc_0061134A: call [0040104Ch] ; __vbaFreeObjList
  loc_00611350: mov eax, var_C4
  loc_00611356: add esp, 00000014h
  loc_00611359: lea ecx, var_24
  loc_0061135C: mov edx, [eax]
  loc_0061135E: push ecx
  loc_0061135F: push eax
  loc_00611360: call [edx+0000008Ch]
  loc_00611366: test eax, eax
  loc_00611368: fnclex
  loc_0061136A: jge 00611380h
  loc_0061136C: mov edx, var_C4
  loc_00611372: push 0000008Ch
  loc_00611377: push 0043B5D0h
  loc_0061137C: push edx
  loc_0061137D: push eax
  loc_0061137E: call edi
  loc_00611380: lea ebx, var_28
  loc_00611383: mov eax, var_24
  loc_00611386: push ebx
  loc_00611387: mov edx, 00000008h
  loc_0061138C: sub esp, 00000010h
  loc_0061138F: mov var_70, edx
  loc_00611392: mov ebx, esp
  loc_00611394: mov ecx, 0043B848h ; "Section4"
  loc_00611399: mov var_68, ecx
  loc_0061139C: mov esi, [eax]
  loc_0061139E: mov [ebx], edx
  loc_006113A0: mov edx, var_6C
  loc_006113A3: push eax
  loc_006113A4: mov var_AC, eax
  loc_006113AA: mov [ebx+00000004h], edx
  loc_006113AD: mov [ebx+00000008h], ecx
  loc_006113B0: mov ecx, var_64
  loc_006113B3: mov [ebx+0000000Ch], ecx
  loc_006113B6: call [esi+0000001Ch]
  loc_006113B9: test eax, eax
  loc_006113BB: fnclex
  loc_006113BD: jge 006113D0h
  loc_006113BF: mov edx, var_AC
  loc_006113C5: push 0000001Ch
  loc_006113C7: push 00439DD8h
  loc_006113CC: push edx
  loc_006113CD: push eax
  loc_006113CE: call edi
  loc_006113D0: mov eax, var_28
  loc_006113D3: lea edx, var_2C
  loc_006113D6: push edx
  loc_006113D7: push eax
  loc_006113D8: mov ecx, [eax]
  loc_006113DA: mov esi, eax
  loc_006113DC: call [ecx+0000003Ch]
  loc_006113DF: test eax, eax
  loc_006113E1: fnclex
  loc_006113E3: jge 006113F0h
  loc_006113E5: push 0000003Ch
  loc_006113E7: push 00439BF8h
  loc_006113EC: push esi
  loc_006113ED: push eax
  loc_006113EE: call edi
  loc_006113F0: lea ebx, var_30
  loc_006113F3: mov eax, var_2C
  loc_006113F6: push ebx
  loc_006113F7: mov edx, 00000008h
  loc_006113FC: sub esp, 00000010h
  loc_006113FF: mov esi, [eax]
  loc_00611401: mov ebx, esp
  loc_00611403: mov ecx, 0043B888h ; "LabelAlamat"
  loc_00611408: push eax
  loc_00611409: mov var_BC, eax
  loc_0061140F: mov [ebx], edx
  loc_00611411: mov edx, var_7C
  loc_00611414: mov [ebx+00000004h], edx
  loc_00611417: mov [ebx+00000008h], ecx
  loc_0061141A: mov ecx, var_74
  loc_0061141D: mov [ebx+0000000Ch], ecx
  loc_00611420: call [esi+0000001Ch]
  loc_00611423: test eax, eax
  loc_00611425: fnclex
  loc_00611427: jge 0061143Ah
  loc_00611429: mov edx, var_BC
  loc_0061142F: push 0000001Ch
  loc_00611431: push 0043B7D8h
  loc_00611436: push edx
  loc_00611437: push eax
  loc_00611438: call edi
  loc_0061143A: mov ecx, [006680E8h]
  loc_00611440: mov edx, [006680ECh]
  loc_00611446: sub esp, 00000010h
  loc_00611449: mov eax, esp
  loc_0061144B: push 0043B85Ch ; "Caption"
  loc_00611450: mov [eax], ecx
  loc_00611452: mov ecx, [006680F0h]
  loc_00611458: mov [eax+00000004h], edx
  loc_0061145B: mov edx, [006680F4h]
  loc_00611461: mov [eax+00000008h], ecx
  loc_00611464: mov [eax+0000000Ch], edx
  loc_00611467: mov eax, var_30
  loc_0061146A: push eax
  loc_0061146B: call [00401094h] ; __vbaLateMemSt
  loc_00611471: lea ecx, var_30
  loc_00611474: lea edx, var_2C
  loc_00611477: push ecx
  loc_00611478: lea eax, var_28
  loc_0061147B: push edx
  loc_0061147C: lea ecx, var_24
  loc_0061147F: push eax
  loc_00611480: push ecx
  loc_00611481: push 00000004h
  loc_00611483: call [0040104Ch] ; __vbaFreeObjList
  loc_00611489: mov eax, var_C4
  loc_0061148F: add esp, 00000014h
  loc_00611492: lea ecx, var_24
  loc_00611495: mov edx, [eax]
  loc_00611497: push ecx
  loc_00611498: push eax
  loc_00611499: call [edx+0000008Ch]
  loc_0061149F: test eax, eax
  loc_006114A1: fnclex
  loc_006114A3: jge 006114B9h
  loc_006114A5: mov edx, var_C4
  loc_006114AB: push 0000008Ch
  loc_006114B0: push 0043B5D0h
  loc_006114B5: push edx
  loc_006114B6: push eax
  loc_006114B7: call edi
  loc_006114B9: lea ebx, var_28
  loc_006114BC: mov eax, var_24
  loc_006114BF: push ebx
  loc_006114C0: mov edx, 00000008h
  loc_006114C5: sub esp, 00000010h
  loc_006114C8: mov var_70, edx
  loc_006114CB: mov ebx, esp
  loc_006114CD: mov ecx, 0043B848h ; "Section4"
  loc_006114D2: mov var_68, ecx
  loc_006114D5: mov esi, [eax]
  loc_006114D7: mov [ebx], edx
  loc_006114D9: mov edx, var_6C
  loc_006114DC: push eax
  loc_006114DD: mov var_AC, eax
  loc_006114E3: mov [ebx+00000004h], edx
  loc_006114E6: mov [ebx+00000008h], ecx
  loc_006114E9: mov ecx, var_64
  loc_006114EC: mov [ebx+0000000Ch], ecx
  loc_006114EF: call [esi+0000001Ch]
  loc_006114F2: test eax, eax
  loc_006114F4: fnclex
  loc_006114F6: jge 00611509h
  loc_006114F8: mov edx, var_AC
  loc_006114FE: push 0000001Ch
  loc_00611500: push 00439DD8h
  loc_00611505: push edx
  loc_00611506: push eax
  loc_00611507: call edi
  loc_00611509: mov eax, var_28
  loc_0061150C: lea edx, var_2C
  loc_0061150F: push edx
  loc_00611510: push eax
  loc_00611511: mov ecx, [eax]
  loc_00611513: mov esi, eax
  loc_00611515: call [ecx+0000003Ch]
  loc_00611518: test eax, eax
  loc_0061151A: fnclex
  loc_0061151C: jge 00611529h
  loc_0061151E: push 0000003Ch
  loc_00611520: push 00439BF8h
  loc_00611525: push esi
  loc_00611526: push eax
  loc_00611527: call edi
  loc_00611529: lea ebx, var_30
  loc_0061152C: mov eax, var_2C
  loc_0061152F: push ebx
  loc_00611530: mov edx, 00000008h
  loc_00611535: sub esp, 00000010h
  loc_00611538: mov esi, [eax]
  loc_0061153A: mov ebx, esp
  loc_0061153C: mov ecx, 0043B8A4h ; "LabelKota"
  loc_00611541: push eax
  loc_00611542: mov var_BC, eax
  loc_00611548: mov [ebx], edx
  loc_0061154A: mov edx, var_7C
  loc_0061154D: mov [ebx+00000004h], edx
  loc_00611550: mov [ebx+00000008h], ecx
  loc_00611553: mov ecx, var_74
  loc_00611556: mov [ebx+0000000Ch], ecx
  loc_00611559: call [esi+0000001Ch]
  loc_0061155C: test eax, eax
  loc_0061155E: fnclex
  loc_00611560: jge 00611573h
  loc_00611562: mov edx, var_BC
  loc_00611568: push 0000001Ch
  loc_0061156A: push 0043B7D8h
  loc_0061156F: push edx
  loc_00611570: push eax
  loc_00611571: call edi
  loc_00611573: mov ecx, [006680F8h]
  loc_00611579: mov edx, [006680FCh]
  loc_0061157F: sub esp, 00000010h
  loc_00611582: mov eax, esp
  loc_00611584: push 0043B85Ch ; "Caption"
  loc_00611589: mov [eax], ecx
  loc_0061158B: mov ecx, [00668100h]
  loc_00611591: mov [eax+00000004h], edx
  loc_00611594: mov edx, [00668104h]
  loc_0061159A: mov [eax+00000008h], ecx
  loc_0061159D: mov [eax+0000000Ch], edx
  loc_006115A0: mov eax, var_30
  loc_006115A3: push eax
  loc_006115A4: call [00401094h] ; __vbaLateMemSt
  loc_006115AA: lea ecx, var_30
  loc_006115AD: lea edx, var_2C
  loc_006115B0: push ecx
  loc_006115B1: lea eax, var_28
  loc_006115B4: push edx
  loc_006115B5: lea ecx, var_24
  loc_006115B8: push eax
  loc_006115B9: push ecx
  loc_006115BA: push 00000004h
  loc_006115BC: call [0040104Ch] ; __vbaFreeObjList
  loc_006115C2: mov eax, var_C4
  loc_006115C8: add esp, 00000014h
  loc_006115CB: lea ecx, var_24
  loc_006115CE: mov edx, [eax]
  loc_006115D0: push ecx
  loc_006115D1: push eax
  loc_006115D2: call [edx+0000008Ch]
  loc_006115D8: test eax, eax
  loc_006115DA: fnclex
  loc_006115DC: jge 006115F2h
  loc_006115DE: mov edx, var_C4
  loc_006115E4: push 0000008Ch
  loc_006115E9: push 0043B5D0h
  loc_006115EE: push edx
  loc_006115EF: push eax
  loc_006115F0: call edi
  loc_006115F2: lea ebx, var_28
  loc_006115F5: mov eax, var_24
  loc_006115F8: push ebx
  loc_006115F9: mov edx, 00000008h
  loc_006115FE: sub esp, 00000010h
  loc_00611601: mov var_70, edx
  loc_00611604: mov ebx, esp
  loc_00611606: mov ecx, 0043B848h ; "Section4"
  loc_0061160B: mov var_68, ecx
  loc_0061160E: mov esi, [eax]
  loc_00611610: mov [ebx], edx
  loc_00611612: mov edx, var_6C
  loc_00611615: push eax
  loc_00611616: mov var_AC, eax
  loc_0061161C: mov [ebx+00000004h], edx
  loc_0061161F: mov [ebx+00000008h], ecx
  loc_00611622: mov ecx, var_64
  loc_00611625: mov [ebx+0000000Ch], ecx
  loc_00611628: call [esi+0000001Ch]
  loc_0061162B: test eax, eax
  loc_0061162D: fnclex
  loc_0061162F: jge 00611642h
  loc_00611631: mov edx, var_AC
  loc_00611637: push 0000001Ch
  loc_00611639: push 00439DD8h
  loc_0061163E: push edx
  loc_0061163F: push eax
  loc_00611640: call edi
  loc_00611642: mov eax, var_28
  loc_00611645: lea edx, var_2C
  loc_00611648: push edx
  loc_00611649: push eax
  loc_0061164A: mov ecx, [eax]
  loc_0061164C: mov esi, eax
  loc_0061164E: call [ecx+0000003Ch]
  loc_00611651: test eax, eax
  loc_00611653: fnclex
  loc_00611655: jge 00611662h
  loc_00611657: push 0000003Ch
  loc_00611659: push 00439BF8h
  loc_0061165E: push esi
  loc_0061165F: push eax
  loc_00611660: call edi
  loc_00611662: lea ebx, var_30
  loc_00611665: mov eax, var_2C
  loc_00611668: push ebx
  loc_00611669: mov edx, 00000008h
  loc_0061166E: sub esp, 00000010h
  loc_00611671: mov esi, [eax]
  loc_00611673: mov ebx, esp
  loc_00611675: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0061167A: push eax
  loc_0061167B: mov var_BC, eax
  loc_00611681: mov [ebx], edx
  loc_00611683: mov edx, var_7C
  loc_00611686: mov [ebx+00000004h], edx
  loc_00611689: mov [ebx+00000008h], ecx
  loc_0061168C: mov ecx, var_74
  loc_0061168F: mov [ebx+0000000Ch], ecx
  loc_00611692: call [esi+0000001Ch]
  loc_00611695: test eax, eax
  loc_00611697: fnclex
  loc_00611699: jge 006116ACh
  loc_0061169B: mov edx, var_BC
  loc_006116A1: push 0000001Ch
  loc_006116A3: push 0043B7D8h
  loc_006116A8: push edx
  loc_006116A9: push eax
  loc_006116AA: call edi
  loc_006116AC: mov ecx, [00668108h]
  loc_006116B2: mov edx, [0066810Ch]
  loc_006116B8: sub esp, 00000010h
  loc_006116BB: mov eax, esp
  loc_006116BD: push 0043B85Ch ; "Caption"
  loc_006116C2: mov [eax], ecx
  loc_006116C4: mov ecx, [00668110h]
  loc_006116CA: mov [eax+00000004h], edx
  loc_006116CD: mov edx, [00668114h]
  loc_006116D3: mov [eax+00000008h], ecx
  loc_006116D6: mov [eax+0000000Ch], edx
  loc_006116D9: mov eax, var_30
  loc_006116DC: push eax
  loc_006116DD: call [00401094h] ; __vbaLateMemSt
  loc_006116E3: lea ecx, var_30
  loc_006116E6: lea edx, var_2C
  loc_006116E9: push ecx
  loc_006116EA: lea eax, var_28
  loc_006116ED: push edx
  loc_006116EE: lea ecx, var_24
  loc_006116F1: push eax
  loc_006116F2: push ecx
  loc_006116F3: push 00000004h
  loc_006116F5: call [0040104Ch] ; __vbaFreeObjList
  loc_006116FB: mov eax, var_C4
  loc_00611701: add esp, 00000014h
  loc_00611704: mov edx, [eax]
  loc_00611706: push 0000000Ah
  loc_00611708: push eax
  loc_00611709: call [edx+00000048h]
  loc_0061170C: test eax, eax
  loc_0061170E: fnclex
  loc_00611710: jge 00611723h
  loc_00611712: mov ecx, var_C4
  loc_00611718: push 00000048h
  loc_0061171A: push 0043B5D0h
  loc_0061171F: push ecx
  loc_00611720: push eax
  loc_00611721: call edi
  loc_00611723: mov eax, var_C4
  loc_00611729: push 0000000Ah
  loc_0061172B: push eax
  loc_0061172C: mov edx, [eax]
  loc_0061172E: call [edx+00000058h]
  loc_00611731: test eax, eax
  loc_00611733: fnclex
  loc_00611735: jge 00611748h
  loc_00611737: mov ecx, var_C4
  loc_0061173D: push 00000058h
  loc_0061173F: push 0043B5D0h
  loc_00611744: push ecx
  loc_00611745: push eax
  loc_00611746: call edi
  loc_00611748: sub esp, 00000010h
  loc_0061174B: mov eax, 00000002h
  loc_00611750: mov edx, esp
  loc_00611752: mov ecx, eax
  loc_00611754: mov var_70, ecx
  loc_00611757: mov var_68, eax
  loc_0061175A: mov [edx], ecx
  loc_0061175C: mov ecx, var_6C
  loc_0061175F: push 8001000Eh
  loc_00611764: mov [edx+00000004h], ecx
  loc_00611767: mov ecx, var_C4
  loc_0061176D: push ecx
  loc_0061176E: mov [edx+00000008h], eax
  loc_00611771: mov eax, var_64
  loc_00611774: mov [edx+0000000Ch], eax
  loc_00611777: call [00401280h] ; __vbaLateIdSt
  loc_0061177D: mov edx, var_C4
  loc_00611783: push 00000000h
  loc_00611785: push 80011003h
  loc_0061178A: push edx
  loc_0061178B: call [0040102Ch] ; __vbaLateIdCall
  loc_00611791: add esp, 0000000Ch
  loc_00611794: lea eax, var_C4
  loc_0061179A: push 00000000h
  loc_0061179C: push eax
  loc_0061179D: call [004010C8h] ; __vbaObjSetAddref
  loc_006117A3: push 0061180Dh
  loc_006117A8: jmp 006117E3h
  loc_006117AA: lea ecx, var_20
  loc_006117AD: call [0040129Ch] ; __vbaFreeStr
  loc_006117B3: lea ecx, var_30
  loc_006117B6: lea edx, var_2C
  loc_006117B9: push ecx
  loc_006117BA: lea eax, var_28
  loc_006117BD: push edx
  loc_006117BE: lea ecx, var_24
  loc_006117C1: push eax
  loc_006117C2: push ecx
  loc_006117C3: push 00000004h
  loc_006117C5: call [0040104Ch] ; __vbaFreeObjList
  loc_006117CB: lea edx, var_60
  loc_006117CE: lea eax, var_50
  loc_006117D1: push edx
  loc_006117D2: lea ecx, var_40
  loc_006117D5: push eax
  loc_006117D6: push ecx
  loc_006117D7: push 00000003h
  loc_006117D9: call [0040103Ch] ; __vbaFreeVarList
  loc_006117DF: add esp, 00000024h
  loc_006117E2: ret
  loc_006117E3: lea edx, var_C8
  loc_006117E9: lea eax, var_C4
  loc_006117EF: push edx
  loc_006117F0: push eax
  loc_006117F1: push 00000002h
  loc_006117F3: call [0040104Ch] ; __vbaFreeObjList
  loc_006117F9: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_006117FF: add esp, 0000000Ch
  loc_00611802: lea ecx, var_18
  loc_00611805: call __vbaFreeStr
  loc_00611807: lea ecx, var_1C
  loc_0061180A: call __vbaFreeStr
  loc_0061180C: ret
  loc_0061180D: mov ecx, var_10
  loc_00611810: pop edi
  loc_00611811: pop esi
  loc_00611812: mov fs:[00000000h], ecx
  loc_00611819: pop ebx
  loc_0061181A: mov esp, ebp
  loc_0061181C: pop ebp
  loc_0061181D: ret
End Sub

Private Sub Proc_34_9_611830
  loc_00611830: push ebp
  loc_00611831: mov ebp, esp
  loc_00611833: sub esp, 00000008h
  loc_00611836: push 00405696h ; __vbaExceptHandler
  loc_0061183B: mov eax, fs:[00000000h]
  loc_00611841: push eax
  loc_00611842: mov fs:[00000000h], esp
  loc_00611849: sub esp, 000000D0h
  loc_0061184F: push ebx
  loc_00611850: push esi
  loc_00611851: push edi
  loc_00611852: mov var_8, esp
  loc_00611855: mov var_4, 00404530h
  loc_0061185C: xor ebx, ebx
  loc_0061185E: mov var_18, ebx
  loc_00611861: mov var_1C, ebx
  loc_00611864: mov var_20, ebx
  loc_00611867: mov var_24, ebx
  loc_0061186A: mov var_28, ebx
  loc_0061186D: mov var_2C, ebx
  loc_00611870: mov var_30, ebx
  loc_00611873: mov var_40, ebx
  loc_00611876: mov var_50, ebx
  loc_00611879: mov var_60, ebx
  loc_0061187C: mov var_70, ebx
  loc_0061187F: mov var_80, ebx
  loc_00611882: mov var_90, ebx
  loc_00611888: mov var_A4, ebx
  loc_0061188E: mov var_C4, ebx
  loc_00611894: mov var_C8, ebx
  loc_0061189A: call 0055B320h
  loc_0061189F: mov edx, 0042DFECh
  loc_006118A4: lea ecx, var_18
  loc_006118A7: call [004011FCh] ; __vbaStrCopy
  loc_006118AD: push 0043C1BCh ; " SELECT No_Retur_jual, no_nota, tgl_retur_jual, nama_pelanggan,"
  loc_006118B2: push 0043C240h ; "tot_retur_jual FROM Retur_jual order BY No_Retur_jual"
  loc_006118B7: call [00401070h] ; __vbaStrCat
  loc_006118BD: mov edx, eax
  loc_006118BF: lea ecx, var_18
  loc_006118C2: call [0040126Ch] ; __vbaStrMove
  loc_006118C8: push 0042DF30h
  loc_006118CD: call [00401168h] ; __vbaNew
  loc_006118D3: push eax
  loc_006118D4: push 00668090h
  loc_006118D9: call [004010B8h] ; __vbaObjSet
  loc_006118DF: cmp [0066803Ch], ebx
  loc_006118E5: jnz 006118F7h
  loc_006118E7: push 0066803Ch
  loc_006118EC: push 0042DEFCh
  loc_006118F1: call [004011E8h] ; __vbaNew2
  loc_006118F7: mov esi, [0066803Ch]
  loc_006118FD: lea ecx, var_40
  loc_00611900: mov var_38, 80020004h
  loc_00611907: mov var_40, 0000000Ah
  loc_0061190E: call [0040123Ch] ; __vbaFreeVarg
  loc_00611914: mov eax, [esi]
  loc_00611916: lea ecx, var_24
  loc_00611919: push ecx
  loc_0061191A: mov ecx, var_18
  loc_0061191D: lea edx, var_40
  loc_00611920: push 00000001h
  loc_00611922: push edx
  loc_00611923: push ecx
  loc_00611924: push esi
  loc_00611925: call [eax+00000040h]
  loc_00611928: cmp eax, ebx
  loc_0061192A: fnclex
  loc_0061192C: jge 00611941h
  loc_0061192E: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00611934: push 00000040h
  loc_00611936: push 0042E1B0h
  loc_0061193B: push esi
  loc_0061193C: push eax
  loc_0061193D: call edi
  loc_0061193F: jmp 00611947h
  loc_00611941: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00611947: mov edx, var_24
  loc_0061194A: mov esi, [00401268h] ; __vbaCastObj
  loc_00611950: push 0042DF20h
  loc_00611955: push edx
  loc_00611956: call __vbaCastObj
  loc_00611958: push eax
  loc_00611959: push 00668090h
  loc_0061195E: call [004010B8h] ; __vbaObjSet
  loc_00611964: lea ecx, var_24
  loc_00611967: call [004012A0h] ; __vbaFreeObj
  loc_0061196D: lea ecx, var_40
  loc_00611970: call [00401020h] ; __vbaFreeVar
  loc_00611976: cmp [006685A8h], ebx
  loc_0061197C: jnz 0061198Eh
  loc_0061197E: push 006685A8h
  loc_00611983: push 00407EC8h
  loc_00611988: call [004011E8h] ; __vbaNew2
  loc_0061198E: mov eax, [006685A8h]
  loc_00611993: lea ecx, var_C4
  loc_00611999: push eax
  loc_0061199A: push ecx
  loc_0061199B: call [004010C8h] ; __vbaObjSetAddref
  loc_006119A1: mov edx, var_C4
  loc_006119A7: push 0043B7B8h
  loc_006119AC: push ebx
  loc_006119AD: mov edx, [edx]
  loc_006119AF: mov var_E0, edx
  loc_006119B5: call __vbaCastObj
  loc_006119B7: push eax
  loc_006119B8: lea eax, var_24
  loc_006119BB: push eax
  loc_006119BC: call [004010B8h] ; __vbaObjSet
  loc_006119C2: mov ecx, var_C4
  loc_006119C8: mov edx, var_E0
  loc_006119CE: push eax
  loc_006119CF: push ecx
  loc_006119D0: call [edx+00000028h]
  loc_006119D3: cmp eax, ebx
  loc_006119D5: fnclex
  loc_006119D7: jge 006119EAh
  loc_006119D9: mov ecx, var_C4
  loc_006119DF: push 00000028h
  loc_006119E1: push 0043B5D0h
  loc_006119E6: push ecx
  loc_006119E7: push eax
  loc_006119E8: call edi
  loc_006119EA: lea ecx, var_24
  loc_006119ED: call [004012A0h] ; __vbaFreeObj
  loc_006119F3: mov eax, var_C4
  loc_006119F9: push 0042DFECh
  loc_006119FE: push eax
  loc_006119FF: mov edx, [eax]
  loc_00611A01: call [edx+00000030h]
  loc_00611A04: cmp eax, ebx
  loc_00611A06: fnclex
  loc_00611A08: jge 00611A1Bh
  loc_00611A0A: mov ecx, var_C4
  loc_00611A10: push 00000030h
  loc_00611A12: push 0043B5D0h
  loc_00611A17: push ecx
  loc_00611A18: push eax
  loc_00611A19: call edi
  loc_00611A1B: mov eax, [00668090h]
  loc_00611A20: lea ecx, var_24
  loc_00611A23: push ecx
  loc_00611A24: push eax
  loc_00611A25: mov edx, [eax]
  loc_00611A27: call [edx+00000114h]
  loc_00611A2D: cmp eax, ebx
  loc_00611A2F: fnclex
  loc_00611A31: jge 00611A47h
  loc_00611A33: mov edx, [00668090h]
  loc_00611A39: push 00000114h
  loc_00611A3E: push 0043B7C8h
  loc_00611A43: push edx
  loc_00611A44: push eax
  loc_00611A45: call edi
  loc_00611A47: mov eax, var_C4
  loc_00611A4D: mov ecx, var_24
  loc_00611A50: push 0043B7B8h
  loc_00611A55: push ecx
  loc_00611A56: mov esi, [eax]
  loc_00611A58: call [00401268h] ; __vbaCastObj
  loc_00611A5E: lea edx, var_28
  loc_00611A61: push eax
  loc_00611A62: push edx
  loc_00611A63: call [004010B8h] ; __vbaObjSet
  loc_00611A69: push eax
  loc_00611A6A: mov eax, var_C4
  loc_00611A70: push eax
  loc_00611A71: call [esi+00000028h]
  loc_00611A74: cmp eax, ebx
  loc_00611A76: fnclex
  loc_00611A78: jge 00611A8Bh
  loc_00611A7A: mov ecx, var_C4
  loc_00611A80: push 00000028h
  loc_00611A82: push 0043B5D0h
  loc_00611A87: push ecx
  loc_00611A88: push eax
  loc_00611A89: call edi
  loc_00611A8B: lea edx, var_28
  loc_00611A8E: lea eax, var_24
  loc_00611A91: push edx
  loc_00611A92: push eax
  loc_00611A93: push 00000002h
  loc_00611A95: call [0040104Ch] ; __vbaFreeObjList
  loc_00611A9B: mov eax, var_C4
  loc_00611AA1: add esp, 0000000Ch
  loc_00611AA4: lea edx, var_24
  loc_00611AA7: mov ecx, [eax]
  loc_00611AA9: push edx
  loc_00611AAA: push eax
  loc_00611AAB: call [ecx+0000008Ch]
  loc_00611AB1: cmp eax, ebx
  loc_00611AB3: fnclex
  loc_00611AB5: jge 00611ACBh
  loc_00611AB7: mov ecx, var_C4
  loc_00611ABD: push 0000008Ch
  loc_00611AC2: push 0043B5D0h
  loc_00611AC7: push ecx
  loc_00611AC8: push eax
  loc_00611AC9: call edi
  loc_00611ACB: lea esi, var_28
  loc_00611ACE: mov eax, var_24
  loc_00611AD1: push esi
  loc_00611AD2: mov ecx, 00000008h
  loc_00611AD7: sub esp, 00000010h
  loc_00611ADA: mov var_70, ecx
  loc_00611ADD: mov esi, esp
  loc_00611ADF: mov var_68, 00439FB0h ; "Section1"
  loc_00611AE6: mov edx, [eax]
  loc_00611AE8: push eax
  loc_00611AE9: mov [esi], ecx
  loc_00611AEB: mov ecx, var_6C
  loc_00611AEE: mov var_AC, eax
  loc_00611AF4: mov [esi+00000004h], ecx
  loc_00611AF7: mov ecx, var_68
  loc_00611AFA: mov [esi+00000008h], ecx
  loc_00611AFD: mov ecx, var_64
  loc_00611B00: mov [esi+0000000Ch], ecx
  loc_00611B03: call [edx+0000001Ch]
  loc_00611B06: cmp eax, ebx
  loc_00611B08: fnclex
  loc_00611B0A: jge 00611B1Dh
  loc_00611B0C: mov edx, var_AC
  loc_00611B12: push 0000001Ch
  loc_00611B14: push 00439DD8h
  loc_00611B19: push edx
  loc_00611B1A: push eax
  loc_00611B1B: call edi
  loc_00611B1D: mov eax, var_28
  loc_00611B20: lea edx, var_2C
  loc_00611B23: push edx
  loc_00611B24: push eax
  loc_00611B25: mov ecx, [eax]
  loc_00611B27: mov esi, eax
  loc_00611B29: call [ecx+0000003Ch]
  loc_00611B2C: cmp eax, ebx
  loc_00611B2E: fnclex
  loc_00611B30: jge 00611B3Dh
  loc_00611B32: push 0000003Ch
  loc_00611B34: push 00439BF8h
  loc_00611B39: push esi
  loc_00611B3A: push eax
  loc_00611B3B: call edi
  loc_00611B3D: mov eax, var_2C
  loc_00611B40: mov var_2C, ebx
  loc_00611B43: push eax
  loc_00611B44: lea eax, var_C8
  loc_00611B4A: push eax
  loc_00611B4B: call [004010B8h] ; __vbaObjSet
  loc_00611B51: lea ecx, var_28
  loc_00611B54: lea edx, var_24
  loc_00611B57: push ecx
  loc_00611B58: push edx
  loc_00611B59: push 00000002h
  loc_00611B5B: call [0040104Ch] ; __vbaFreeObjList
  loc_00611B61: mov eax, var_C8
  loc_00611B67: add esp, 0000000Ch
  loc_00611B6A: lea edx, var_A4
  loc_00611B70: mov ecx, [eax]
  loc_00611B72: push edx
  loc_00611B73: push eax
  loc_00611B74: call [ecx+00000020h]
  loc_00611B77: cmp eax, ebx
  loc_00611B79: fnclex
  loc_00611B7B: jge 00611B8Eh
  loc_00611B7D: mov ecx, var_C8
  loc_00611B83: push 00000020h
  loc_00611B85: push 0043B7D8h
  loc_00611B8A: push ecx
  loc_00611B8B: push eax
  loc_00611B8C: call edi
  loc_00611B8E: mov ecx, var_A4
  loc_00611B94: call [00401138h] ; __vbaI2I4
  loc_00611B9A: mov var_D0, eax
  loc_00611BA0: mov ebx, 00000001h
  loc_00611BA5: cmp bx, var_D0
  loc_00611BAC: mov var_14, ebx
  loc_00611BAF: jg 00611E20h
  loc_00611BB5: lea esi, var_24
  loc_00611BB8: mov ecx, var_C8
  loc_00611BBE: push esi
  loc_00611BBF: mov eax, 00000002h
  loc_00611BC4: sub esp, 00000010h
  loc_00611BC7: mov var_70, eax
  loc_00611BCA: mov esi, esp
  loc_00611BCC: mov var_68, bx
  loc_00611BD0: mov edx, [ecx]
  loc_00611BD2: push ecx
  loc_00611BD3: mov [esi], eax
  loc_00611BD5: mov eax, var_6C
  loc_00611BD8: mov [esi+00000004h], eax
  loc_00611BDB: mov eax, var_68
  loc_00611BDE: mov [esi+00000008h], eax
  loc_00611BE1: mov eax, var_64
  loc_00611BE4: mov [esi+0000000Ch], eax
  loc_00611BE7: call [edx+0000001Ch]
  loc_00611BEA: test eax, eax
  loc_00611BEC: fnclex
  loc_00611BEE: jge 00611C01h
  loc_00611BF0: mov ecx, var_C8
  loc_00611BF6: push 0000001Ch
  loc_00611BF8: push 0043B7D8h
  loc_00611BFD: push ecx
  loc_00611BFE: push eax
  loc_00611BFF: call edi
  loc_00611C01: mov edx, var_24
  loc_00611C04: push 0043B7E8h
  loc_00611C09: push edx
  loc_00611C0A: call [004011CCh] ; __vbaCheckType
  loc_00611C10: lea ecx, var_24
  loc_00611C13: mov si, ax
  loc_00611C16: call [004012A0h] ; __vbaFreeObj
  loc_00611C1C: test si, si
  loc_00611C1F: Unknown_795()
  loc_00611C25: mov var_68, bx
  loc_00611C29: lea ebx, var_24
  loc_00611C2C: push ebx
  loc_00611C2D: mov ecx, var_C8
  loc_00611C33: sub esp, 00000010h
  loc_00611C36: mov eax, 00000002h
  loc_00611C3B: mov ebx, esp
  loc_00611C3D: mov var_70, eax
  loc_00611C40: mov edx, [ecx]
  loc_00611C42: mov esi, 0042DFECh
  loc_00611C47: mov [ebx], eax
  loc_00611C49: mov eax, var_6C
  loc_00611C4C: push ecx
  loc_00611C4D: mov var_78, esi
  loc_00611C50: mov [ebx+00000004h], eax
  loc_00611C53: mov eax, var_68
  loc_00611C56: mov edi, 00000008h
  loc_00611C5B: mov [ebx+00000008h], eax
  loc_00611C5E: mov eax, var_64
  loc_00611C61: mov [ebx+0000000Ch], eax
  loc_00611C64: call [edx+0000001Ch]
  loc_00611C67: test eax, eax
  loc_00611C69: fnclex
  loc_00611C6B: jge 00611C82h
  loc_00611C6D: mov ecx, var_C8
  loc_00611C73: push 0000001Ch
  loc_00611C75: push 0043B7D8h
  loc_00611C7A: push ecx
  loc_00611C7B: push eax
  loc_00611C7C: call [00401080h] ; __vbaHresultCheckObj
  loc_00611C82: mov eax, var_7C
  loc_00611C85: sub esp, 00000010h
  loc_00611C88: mov edx, esp
  loc_00611C8A: mov ecx, var_74
  loc_00611C8D: push 0043B7F8h ; "DataMember"
  loc_00611C92: mov [edx], edi
  loc_00611C94: mov [edx+00000004h], eax
  loc_00611C97: mov [edx+00000008h], esi
  loc_00611C9A: mov [edx+0000000Ch], ecx
  loc_00611C9D: mov edx, var_24
  loc_00611CA0: push edx
  loc_00611CA1: call [00401094h] ; __vbaLateMemSt
  loc_00611CA7: lea ecx, var_24
  loc_00611CAA: call [004012A0h] ; __vbaFreeObj
  loc_00611CB0: mov eax, [00668090h]
  loc_00611CB5: lea edx, var_24
  loc_00611CB8: push edx
  loc_00611CB9: push eax
  loc_00611CBA: mov ecx, [eax]
  loc_00611CBC: call [ecx+00000054h]
  loc_00611CBF: test eax, eax
  loc_00611CC1: fnclex
  loc_00611CC3: jge 00611CDAh
  loc_00611CC5: mov ecx, [00668090h]
  loc_00611CCB: push 00000054h
  loc_00611CCD: push 0042DF44h
  loc_00611CD2: push ecx
  loc_00611CD3: push eax
  loc_00611CD4: call [00401080h] ; __vbaHresultCheckObj
  loc_00611CDA: mov ebx, var_14
  loc_00611CDD: lea edi, var_28
  loc_00611CE0: mov dx, bx
  loc_00611CE3: push edi
  loc_00611CE4: sub dx, 0001h
  loc_00611CE8: mov eax, var_24
  loc_00611CEB: jo 006125AEh
  loc_00611CF1: sub esp, 00000010h
  loc_00611CF4: mov ecx, 00000002h
  loc_00611CF9: mov edi, esp
  loc_00611CFB: mov var_70, ecx
  loc_00611CFE: mov var_68, dx
  loc_00611D02: mov edx, [eax]
  loc_00611D04: mov [edi], ecx
  loc_00611D06: mov ecx, var_6C
  loc_00611D09: push eax
  loc_00611D0A: mov esi, eax
  loc_00611D0C: mov [edi+00000004h], ecx
  loc_00611D0F: mov ecx, var_68
  loc_00611D12: mov [edi+00000008h], ecx
  loc_00611D15: mov ecx, var_64
  loc_00611D18: mov [edi+0000000Ch], ecx
  loc_00611D1B: call [edx+00000028h]
  loc_00611D1E: test eax, eax
  loc_00611D20: fnclex
  loc_00611D22: jge 00611D37h
  loc_00611D24: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00611D2A: push 00000028h
  loc_00611D2C: push 0042DFACh
  loc_00611D31: push esi
  loc_00611D32: push eax
  loc_00611D33: call edi
  loc_00611D35: jmp 00611D3Dh
  loc_00611D37: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00611D3D: mov eax, var_28
  loc_00611D40: lea ecx, var_20
  loc_00611D43: push ecx
  loc_00611D44: push eax
  loc_00611D45: mov edx, [eax]
  loc_00611D47: mov esi, eax
  loc_00611D49: call [edx+0000002Ch]
  loc_00611D4C: test eax, eax
  loc_00611D4E: fnclex
  loc_00611D50: jge 00611D5Dh
  loc_00611D52: push 0000002Ch
  loc_00611D54: push 0042DFBCh
  loc_00611D59: push esi
  loc_00611D5A: push eax
  loc_00611D5B: call edi
  loc_00611D5D: mov eax, var_20
  loc_00611D60: lea esi, var_2C
  loc_00611D63: push esi
  loc_00611D64: mov ecx, var_C8
  loc_00611D6A: sub esp, 00000010h
  loc_00611D6D: mov var_38, eax
  loc_00611D70: mov esi, esp
  loc_00611D72: mov eax, 00000002h
  loc_00611D77: mov var_78, bx
  loc_00611D7B: mov var_20, 00000000h
  loc_00611D82: mov [esi], eax
  loc_00611D84: mov eax, var_7C
  loc_00611D87: mov var_40, 00000008h
  loc_00611D8E: mov edx, [ecx]
  loc_00611D90: mov [esi+00000004h], eax
  loc_00611D93: mov eax, var_78
  loc_00611D96: push ecx
  loc_00611D97: mov [esi+00000008h], eax
  loc_00611D9A: mov eax, var_74
  loc_00611D9D: mov [esi+0000000Ch], eax
  loc_00611DA0: call [edx+0000001Ch]
  loc_00611DA3: test eax, eax
  loc_00611DA5: fnclex
  loc_00611DA7: jge 00611DBAh
  loc_00611DA9: mov ecx, var_C8
  loc_00611DAF: push 0000001Ch
  loc_00611DB1: push 0043B7D8h
  loc_00611DB6: push ecx
  loc_00611DB7: push eax
  loc_00611DB8: call edi
  loc_00611DBA: mov eax, var_40
  loc_00611DBD: mov ecx, var_3C
  loc_00611DC0: sub esp, 00000010h
  loc_00611DC3: mov edx, esp
  loc_00611DC5: push 0043B810h ; "DataField"
  loc_00611DCA: mov [edx], eax
  loc_00611DCC: mov eax, var_38
  loc_00611DCF: mov [edx+00000004h], ecx
  loc_00611DD2: mov ecx, var_34
  loc_00611DD5: mov [edx+00000008h], eax
  loc_00611DD8: mov [edx+0000000Ch], ecx
  loc_00611DDB: mov edx, var_2C
  loc_00611DDE: push edx
  loc_00611DDF: call [00401094h] ; __vbaLateMemSt
  loc_00611DE5: lea eax, var_2C
  loc_00611DE8: lea ecx, var_28
  loc_00611DEB: push eax
  loc_00611DEC: lea edx, var_24
  loc_00611DEF: push ecx
  loc_00611DF0: push edx
  loc_00611DF1: push 00000003h
  loc_00611DF3: call [0040104Ch] ; __vbaFreeObjList
  loc_00611DF9: add esp, 00000010h
  loc_00611DFC: lea ecx, var_40
  loc_00611DFF: call [00401020h] ; __vbaFreeVar
  loc_00611E05: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00611E0B: mov eax, 00000001h
  loc_00611E10: add ax, bx
  loc_00611E13: jo 006125AEh
  loc_00611E19: mov ebx, eax
  loc_00611E1B: jmp 00611BA5h
  loc_00611E20: lea eax, var_C8
  loc_00611E26: push 00000000h
  loc_00611E28: push eax
  loc_00611E29: call [004010C8h] ; __vbaObjSetAddref
  loc_00611E2F: lea ecx, var_40
  loc_00611E32: push ecx
  loc_00611E33: call [00401220h] ; rtcGetDateVar
  loc_00611E39: lea edx, var_70
  loc_00611E3C: lea ecx, var_50
  loc_00611E3F: mov var_68, 00433FA8h ; "dd/MM/yyyy"
  loc_00611E46: mov var_70, 00000008h
  loc_00611E4D: call [00401238h] ; __vbaVarDup
  loc_00611E53: push 00000001h
  loc_00611E55: lea edx, var_50
  loc_00611E58: push 00000001h
  loc_00611E5A: lea eax, var_40
  loc_00611E5D: push edx
  loc_00611E5E: lea ecx, var_60
  loc_00611E61: push eax
  loc_00611E62: push ecx
  loc_00611E63: call [00401078h] ; rtcVarFromFormatVar
  loc_00611E69: mov eax, var_C4
  loc_00611E6F: lea ecx, var_24
  loc_00611E72: push ecx
  loc_00611E73: push eax
  loc_00611E74: mov edx, [eax]
  loc_00611E76: call [edx+0000008Ch]
  loc_00611E7C: test eax, eax
  loc_00611E7E: fnclex
  loc_00611E80: jge 00611E96h
  loc_00611E82: mov edx, var_C4
  loc_00611E88: push 0000008Ch
  loc_00611E8D: push 0043B5D0h
  loc_00611E92: push edx
  loc_00611E93: push eax
  loc_00611E94: call edi
  loc_00611E96: lea ebx, var_28
  loc_00611E99: mov eax, var_24
  loc_00611E9C: push ebx
  loc_00611E9D: mov edx, 00000008h
  loc_00611EA2: sub esp, 00000010h
  loc_00611EA5: mov esi, [eax]
  loc_00611EA7: mov ebx, esp
  loc_00611EA9: mov ecx, 0043B848h ; "Section4"
  loc_00611EAE: push eax
  loc_00611EAF: mov var_AC, eax
  loc_00611EB5: mov [ebx], edx
  loc_00611EB7: mov edx, var_7C
  loc_00611EBA: mov [ebx+00000004h], edx
  loc_00611EBD: mov [ebx+00000008h], ecx
  loc_00611EC0: mov ecx, var_74
  loc_00611EC3: mov [ebx+0000000Ch], ecx
  loc_00611EC6: call [esi+0000001Ch]
  loc_00611EC9: test eax, eax
  loc_00611ECB: fnclex
  loc_00611ECD: jge 00611EE0h
  loc_00611ECF: mov edx, var_AC
  loc_00611ED5: push 0000001Ch
  loc_00611ED7: push 00439DD8h
  loc_00611EDC: push edx
  loc_00611EDD: push eax
  loc_00611EDE: call edi
  loc_00611EE0: mov eax, var_28
  loc_00611EE3: lea edx, var_2C
  loc_00611EE6: push edx
  loc_00611EE7: push eax
  loc_00611EE8: mov ecx, [eax]
  loc_00611EEA: mov esi, eax
  loc_00611EEC: call [ecx+0000003Ch]
  loc_00611EEF: test eax, eax
  loc_00611EF1: fnclex
  loc_00611EF3: jge 00611F00h
  loc_00611EF5: push 0000003Ch
  loc_00611EF7: push 00439BF8h
  loc_00611EFC: push esi
  loc_00611EFD: push eax
  loc_00611EFE: call edi
  loc_00611F00: lea ebx, var_30
  loc_00611F03: mov eax, var_2C
  loc_00611F06: push ebx
  loc_00611F07: mov edx, 00000008h
  loc_00611F0C: sub esp, 00000010h
  loc_00611F0F: mov esi, [eax]
  loc_00611F11: mov ebx, esp
  loc_00611F13: mov ecx, 0043B828h ; "labelTanggal"
  loc_00611F18: push eax
  loc_00611F19: mov var_BC, eax
  loc_00611F1F: mov [ebx], edx
  loc_00611F21: mov edx, var_8C
  loc_00611F27: mov [ebx+00000004h], edx
  loc_00611F2A: mov [ebx+00000008h], ecx
  loc_00611F2D: mov ecx, var_84
  loc_00611F33: mov [ebx+0000000Ch], ecx
  loc_00611F36: call [esi+0000001Ch]
  loc_00611F39: test eax, eax
  loc_00611F3B: fnclex
  loc_00611F3D: jge 00611F50h
  loc_00611F3F: mov edx, var_BC
  loc_00611F45: push 0000001Ch
  loc_00611F47: push 0043B7D8h
  loc_00611F4C: push edx
  loc_00611F4D: push eax
  loc_00611F4E: call edi
  loc_00611F50: mov ecx, var_60
  loc_00611F53: mov edx, var_5C
  loc_00611F56: sub esp, 00000010h
  loc_00611F59: mov eax, esp
  loc_00611F5B: push 0043B85Ch ; "Caption"
  loc_00611F60: mov [eax], ecx
  loc_00611F62: mov ecx, var_58
  loc_00611F65: mov [eax+00000004h], edx
  loc_00611F68: mov edx, var_54
  loc_00611F6B: mov [eax+00000008h], ecx
  loc_00611F6E: mov [eax+0000000Ch], edx
  loc_00611F71: mov eax, var_30
  loc_00611F74: push eax
  loc_00611F75: call [00401094h] ; __vbaLateMemSt
  loc_00611F7B: lea ecx, var_30
  loc_00611F7E: lea edx, var_2C
  loc_00611F81: push ecx
  loc_00611F82: lea eax, var_28
  loc_00611F85: push edx
  loc_00611F86: lea ecx, var_24
  loc_00611F89: push eax
  loc_00611F8A: push ecx
  loc_00611F8B: push 00000004h
  loc_00611F8D: call [0040104Ch] ; __vbaFreeObjList
  loc_00611F93: lea edx, var_60
  loc_00611F96: lea eax, var_50
  loc_00611F99: push edx
  loc_00611F9A: lea ecx, var_40
  loc_00611F9D: push eax
  loc_00611F9E: push ecx
  loc_00611F9F: push 00000003h
  loc_00611FA1: call [0040103Ch] ; __vbaFreeVarList
  loc_00611FA7: mov eax, var_C4
  loc_00611FAD: add esp, 00000024h
  loc_00611FB0: lea ecx, var_24
  loc_00611FB3: mov edx, [eax]
  loc_00611FB5: push ecx
  loc_00611FB6: push eax
  loc_00611FB7: call [edx+0000008Ch]
  loc_00611FBD: test eax, eax
  loc_00611FBF: fnclex
  loc_00611FC1: jge 00611FD7h
  loc_00611FC3: mov edx, var_C4
  loc_00611FC9: push 0000008Ch
  loc_00611FCE: push 0043B5D0h
  loc_00611FD3: push edx
  loc_00611FD4: push eax
  loc_00611FD5: call edi
  loc_00611FD7: lea ebx, var_28
  loc_00611FDA: mov eax, var_24
  loc_00611FDD: push ebx
  loc_00611FDE: mov edx, 00000008h
  loc_00611FE3: sub esp, 00000010h
  loc_00611FE6: mov var_70, edx
  loc_00611FE9: mov ebx, esp
  loc_00611FEB: mov ecx, 0043B848h ; "Section4"
  loc_00611FF0: mov var_68, ecx
  loc_00611FF3: mov esi, [eax]
  loc_00611FF5: mov [ebx], edx
  loc_00611FF7: mov edx, var_6C
  loc_00611FFA: push eax
  loc_00611FFB: mov var_AC, eax
  loc_00612001: mov [ebx+00000004h], edx
  loc_00612004: mov [ebx+00000008h], ecx
  loc_00612007: mov ecx, var_64
  loc_0061200A: mov [ebx+0000000Ch], ecx
  loc_0061200D: call [esi+0000001Ch]
  loc_00612010: test eax, eax
  loc_00612012: fnclex
  loc_00612014: jge 00612027h
  loc_00612016: mov edx, var_AC
  loc_0061201C: push 0000001Ch
  loc_0061201E: push 00439DD8h
  loc_00612023: push edx
  loc_00612024: push eax
  loc_00612025: call edi
  loc_00612027: mov eax, var_28
  loc_0061202A: lea edx, var_2C
  loc_0061202D: push edx
  loc_0061202E: push eax
  loc_0061202F: mov ecx, [eax]
  loc_00612031: mov esi, eax
  loc_00612033: call [ecx+0000003Ch]
  loc_00612036: test eax, eax
  loc_00612038: fnclex
  loc_0061203A: jge 00612047h
  loc_0061203C: push 0000003Ch
  loc_0061203E: push 00439BF8h
  loc_00612043: push esi
  loc_00612044: push eax
  loc_00612045: call edi
  loc_00612047: lea ebx, var_30
  loc_0061204A: mov eax, var_2C
  loc_0061204D: push ebx
  loc_0061204E: mov edx, 00000008h
  loc_00612053: sub esp, 00000010h
  loc_00612056: mov esi, [eax]
  loc_00612058: mov ebx, esp
  loc_0061205A: mov ecx, 0043B870h ; "LabelNama"
  loc_0061205F: push eax
  loc_00612060: mov var_BC, eax
  loc_00612066: mov [ebx], edx
  loc_00612068: mov edx, var_7C
  loc_0061206B: mov [ebx+00000004h], edx
  loc_0061206E: mov [ebx+00000008h], ecx
  loc_00612071: mov ecx, var_74
  loc_00612074: mov [ebx+0000000Ch], ecx
  loc_00612077: call [esi+0000001Ch]
  loc_0061207A: test eax, eax
  loc_0061207C: fnclex
  loc_0061207E: jge 00612091h
  loc_00612080: mov edx, var_BC
  loc_00612086: push 0000001Ch
  loc_00612088: push 0043B7D8h
  loc_0061208D: push edx
  loc_0061208E: push eax
  loc_0061208F: call edi
  loc_00612091: mov ecx, [006680D8h]
  loc_00612097: mov edx, [006680DCh]
  loc_0061209D: sub esp, 00000010h
  loc_006120A0: mov eax, esp
  loc_006120A2: push 0043B85Ch ; "Caption"
  loc_006120A7: mov [eax], ecx
  loc_006120A9: mov ecx, [006680E0h]
  loc_006120AF: mov [eax+00000004h], edx
  loc_006120B2: mov edx, [006680E4h]
  loc_006120B8: mov [eax+00000008h], ecx
  loc_006120BB: mov [eax+0000000Ch], edx
  loc_006120BE: mov eax, var_30
  loc_006120C1: push eax
  loc_006120C2: call [00401094h] ; __vbaLateMemSt
  loc_006120C8: lea ecx, var_30
  loc_006120CB: lea edx, var_2C
  loc_006120CE: push ecx
  loc_006120CF: lea eax, var_28
  loc_006120D2: push edx
  loc_006120D3: lea ecx, var_24
  loc_006120D6: push eax
  loc_006120D7: push ecx
  loc_006120D8: push 00000004h
  loc_006120DA: call [0040104Ch] ; __vbaFreeObjList
  loc_006120E0: mov eax, var_C4
  loc_006120E6: add esp, 00000014h
  loc_006120E9: lea ecx, var_24
  loc_006120EC: mov edx, [eax]
  loc_006120EE: push ecx
  loc_006120EF: push eax
  loc_006120F0: call [edx+0000008Ch]
  loc_006120F6: test eax, eax
  loc_006120F8: fnclex
  loc_006120FA: jge 00612110h
  loc_006120FC: mov edx, var_C4
  loc_00612102: push 0000008Ch
  loc_00612107: push 0043B5D0h
  loc_0061210C: push edx
  loc_0061210D: push eax
  loc_0061210E: call edi
  loc_00612110: lea ebx, var_28
  loc_00612113: mov eax, var_24
  loc_00612116: push ebx
  loc_00612117: mov edx, 00000008h
  loc_0061211C: sub esp, 00000010h
  loc_0061211F: mov var_70, edx
  loc_00612122: mov ebx, esp
  loc_00612124: mov ecx, 0043B848h ; "Section4"
  loc_00612129: mov var_68, ecx
  loc_0061212C: mov esi, [eax]
  loc_0061212E: mov [ebx], edx
  loc_00612130: mov edx, var_6C
  loc_00612133: push eax
  loc_00612134: mov var_AC, eax
  loc_0061213A: mov [ebx+00000004h], edx
  loc_0061213D: mov [ebx+00000008h], ecx
  loc_00612140: mov ecx, var_64
  loc_00612143: mov [ebx+0000000Ch], ecx
  loc_00612146: call [esi+0000001Ch]
  loc_00612149: test eax, eax
  loc_0061214B: fnclex
  loc_0061214D: jge 00612160h
  loc_0061214F: mov edx, var_AC
  loc_00612155: push 0000001Ch
  loc_00612157: push 00439DD8h
  loc_0061215C: push edx
  loc_0061215D: push eax
  loc_0061215E: call edi
  loc_00612160: mov eax, var_28
  loc_00612163: lea edx, var_2C
  loc_00612166: push edx
  loc_00612167: push eax
  loc_00612168: mov ecx, [eax]
  loc_0061216A: mov esi, eax
  loc_0061216C: call [ecx+0000003Ch]
  loc_0061216F: test eax, eax
  loc_00612171: fnclex
  loc_00612173: jge 00612180h
  loc_00612175: push 0000003Ch
  loc_00612177: push 00439BF8h
  loc_0061217C: push esi
  loc_0061217D: push eax
  loc_0061217E: call edi
  loc_00612180: lea ebx, var_30
  loc_00612183: mov eax, var_2C
  loc_00612186: push ebx
  loc_00612187: mov edx, 00000008h
  loc_0061218C: sub esp, 00000010h
  loc_0061218F: mov esi, [eax]
  loc_00612191: mov ebx, esp
  loc_00612193: mov ecx, 0043B888h ; "LabelAlamat"
  loc_00612198: push eax
  loc_00612199: mov var_BC, eax
  loc_0061219F: mov [ebx], edx
  loc_006121A1: mov edx, var_7C
  loc_006121A4: mov [ebx+00000004h], edx
  loc_006121A7: mov [ebx+00000008h], ecx
  loc_006121AA: mov ecx, var_74
  loc_006121AD: mov [ebx+0000000Ch], ecx
  loc_006121B0: call [esi+0000001Ch]
  loc_006121B3: test eax, eax
  loc_006121B5: fnclex
  loc_006121B7: jge 006121CAh
  loc_006121B9: mov edx, var_BC
  loc_006121BF: push 0000001Ch
  loc_006121C1: push 0043B7D8h
  loc_006121C6: push edx
  loc_006121C7: push eax
  loc_006121C8: call edi
  loc_006121CA: mov ecx, [006680E8h]
  loc_006121D0: mov edx, [006680ECh]
  loc_006121D6: sub esp, 00000010h
  loc_006121D9: mov eax, esp
  loc_006121DB: push 0043B85Ch ; "Caption"
  loc_006121E0: mov [eax], ecx
  loc_006121E2: mov ecx, [006680F0h]
  loc_006121E8: mov [eax+00000004h], edx
  loc_006121EB: mov edx, [006680F4h]
  loc_006121F1: mov [eax+00000008h], ecx
  loc_006121F4: mov [eax+0000000Ch], edx
  loc_006121F7: mov eax, var_30
  loc_006121FA: push eax
  loc_006121FB: call [00401094h] ; __vbaLateMemSt
  loc_00612201: lea ecx, var_30
  loc_00612204: lea edx, var_2C
  loc_00612207: push ecx
  loc_00612208: lea eax, var_28
  loc_0061220B: push edx
  loc_0061220C: lea ecx, var_24
  loc_0061220F: push eax
  loc_00612210: push ecx
  loc_00612211: push 00000004h
  loc_00612213: call [0040104Ch] ; __vbaFreeObjList
  loc_00612219: mov eax, var_C4
  loc_0061221F: add esp, 00000014h
  loc_00612222: lea ecx, var_24
  loc_00612225: mov edx, [eax]
  loc_00612227: push ecx
  loc_00612228: push eax
  loc_00612229: call [edx+0000008Ch]
  loc_0061222F: test eax, eax
  loc_00612231: fnclex
  loc_00612233: jge 00612249h
  loc_00612235: mov edx, var_C4
  loc_0061223B: push 0000008Ch
  loc_00612240: push 0043B5D0h
  loc_00612245: push edx
  loc_00612246: push eax
  loc_00612247: call edi
  loc_00612249: lea ebx, var_28
  loc_0061224C: mov eax, var_24
  loc_0061224F: push ebx
  loc_00612250: mov edx, 00000008h
  loc_00612255: sub esp, 00000010h
  loc_00612258: mov var_70, edx
  loc_0061225B: mov ebx, esp
  loc_0061225D: mov ecx, 0043B848h ; "Section4"
  loc_00612262: mov var_68, ecx
  loc_00612265: mov esi, [eax]
  loc_00612267: mov [ebx], edx
  loc_00612269: mov edx, var_6C
  loc_0061226C: push eax
  loc_0061226D: mov var_AC, eax
  loc_00612273: mov [ebx+00000004h], edx
  loc_00612276: mov [ebx+00000008h], ecx
  loc_00612279: mov ecx, var_64
  loc_0061227C: mov [ebx+0000000Ch], ecx
  loc_0061227F: call [esi+0000001Ch]
  loc_00612282: test eax, eax
  loc_00612284: fnclex
  loc_00612286: jge 00612299h
  loc_00612288: mov edx, var_AC
  loc_0061228E: push 0000001Ch
  loc_00612290: push 00439DD8h
  loc_00612295: push edx
  loc_00612296: push eax
  loc_00612297: call edi
  loc_00612299: mov eax, var_28
  loc_0061229C: lea edx, var_2C
  loc_0061229F: push edx
  loc_006122A0: push eax
  loc_006122A1: mov ecx, [eax]
  loc_006122A3: mov esi, eax
  loc_006122A5: call [ecx+0000003Ch]
  loc_006122A8: test eax, eax
  loc_006122AA: fnclex
  loc_006122AC: jge 006122B9h
  loc_006122AE: push 0000003Ch
  loc_006122B0: push 00439BF8h
  loc_006122B5: push esi
  loc_006122B6: push eax
  loc_006122B7: call edi
  loc_006122B9: lea ebx, var_30
  loc_006122BC: mov eax, var_2C
  loc_006122BF: push ebx
  loc_006122C0: mov edx, 00000008h
  loc_006122C5: sub esp, 00000010h
  loc_006122C8: mov esi, [eax]
  loc_006122CA: mov ebx, esp
  loc_006122CC: mov ecx, 0043B8A4h ; "LabelKota"
  loc_006122D1: push eax
  loc_006122D2: mov var_BC, eax
  loc_006122D8: mov [ebx], edx
  loc_006122DA: mov edx, var_7C
  loc_006122DD: mov [ebx+00000004h], edx
  loc_006122E0: mov [ebx+00000008h], ecx
  loc_006122E3: mov ecx, var_74
  loc_006122E6: mov [ebx+0000000Ch], ecx
  loc_006122E9: call [esi+0000001Ch]
  loc_006122EC: test eax, eax
  loc_006122EE: fnclex
  loc_006122F0: jge 00612303h
  loc_006122F2: mov edx, var_BC
  loc_006122F8: push 0000001Ch
  loc_006122FA: push 0043B7D8h
  loc_006122FF: push edx
  loc_00612300: push eax
  loc_00612301: call edi
  loc_00612303: mov ecx, [006680F8h]
  loc_00612309: mov edx, [006680FCh]
  loc_0061230F: sub esp, 00000010h
  loc_00612312: mov eax, esp
  loc_00612314: push 0043B85Ch ; "Caption"
  loc_00612319: mov [eax], ecx
  loc_0061231B: mov ecx, [00668100h]
  loc_00612321: mov [eax+00000004h], edx
  loc_00612324: mov edx, [00668104h]
  loc_0061232A: mov [eax+00000008h], ecx
  loc_0061232D: mov [eax+0000000Ch], edx
  loc_00612330: mov eax, var_30
  loc_00612333: push eax
  loc_00612334: call [00401094h] ; __vbaLateMemSt
  loc_0061233A: lea ecx, var_30
  loc_0061233D: lea edx, var_2C
  loc_00612340: push ecx
  loc_00612341: lea eax, var_28
  loc_00612344: push edx
  loc_00612345: lea ecx, var_24
  loc_00612348: push eax
  loc_00612349: push ecx
  loc_0061234A: push 00000004h
  loc_0061234C: call [0040104Ch] ; __vbaFreeObjList
  loc_00612352: mov eax, var_C4
  loc_00612358: add esp, 00000014h
  loc_0061235B: lea ecx, var_24
  loc_0061235E: mov edx, [eax]
  loc_00612360: push ecx
  loc_00612361: push eax
  loc_00612362: call [edx+0000008Ch]
  loc_00612368: test eax, eax
  loc_0061236A: fnclex
  loc_0061236C: jge 00612382h
  loc_0061236E: mov edx, var_C4
  loc_00612374: push 0000008Ch
  loc_00612379: push 0043B5D0h
  loc_0061237E: push edx
  loc_0061237F: push eax
  loc_00612380: call edi
  loc_00612382: lea ebx, var_28
  loc_00612385: mov eax, var_24
  loc_00612388: push ebx
  loc_00612389: mov edx, 00000008h
  loc_0061238E: sub esp, 00000010h
  loc_00612391: mov var_70, edx
  loc_00612394: mov ebx, esp
  loc_00612396: mov ecx, 0043B848h ; "Section4"
  loc_0061239B: mov var_68, ecx
  loc_0061239E: mov esi, [eax]
  loc_006123A0: mov [ebx], edx
  loc_006123A2: mov edx, var_6C
  loc_006123A5: push eax
  loc_006123A6: mov var_AC, eax
  loc_006123AC: mov [ebx+00000004h], edx
  loc_006123AF: mov [ebx+00000008h], ecx
  loc_006123B2: mov ecx, var_64
  loc_006123B5: mov [ebx+0000000Ch], ecx
  loc_006123B8: call [esi+0000001Ch]
  loc_006123BB: test eax, eax
  loc_006123BD: fnclex
  loc_006123BF: jge 006123D2h
  loc_006123C1: mov edx, var_AC
  loc_006123C7: push 0000001Ch
  loc_006123C9: push 00439DD8h
  loc_006123CE: push edx
  loc_006123CF: push eax
  loc_006123D0: call edi
  loc_006123D2: mov eax, var_28
  loc_006123D5: lea edx, var_2C
  loc_006123D8: push edx
  loc_006123D9: push eax
  loc_006123DA: mov ecx, [eax]
  loc_006123DC: mov esi, eax
  loc_006123DE: call [ecx+0000003Ch]
  loc_006123E1: test eax, eax
  loc_006123E3: fnclex
  loc_006123E5: jge 006123F2h
  loc_006123E7: push 0000003Ch
  loc_006123E9: push 00439BF8h
  loc_006123EE: push esi
  loc_006123EF: push eax
  loc_006123F0: call edi
  loc_006123F2: lea ebx, var_30
  loc_006123F5: mov eax, var_2C
  loc_006123F8: push ebx
  loc_006123F9: mov edx, 00000008h
  loc_006123FE: sub esp, 00000010h
  loc_00612401: mov esi, [eax]
  loc_00612403: mov ebx, esp
  loc_00612405: mov ecx, 0043B8BCh ; "LabelTelp"
  loc_0061240A: push eax
  loc_0061240B: mov var_BC, eax
  loc_00612411: mov [ebx], edx
  loc_00612413: mov edx, var_7C
  loc_00612416: mov [ebx+00000004h], edx
  loc_00612419: mov [ebx+00000008h], ecx
  loc_0061241C: mov ecx, var_74
  loc_0061241F: mov [ebx+0000000Ch], ecx
  loc_00612422: call [esi+0000001Ch]
  loc_00612425: test eax, eax
  loc_00612427: fnclex
  loc_00612429: jge 0061243Ch
  loc_0061242B: mov edx, var_BC
  loc_00612431: push 0000001Ch
  loc_00612433: push 0043B7D8h
  loc_00612438: push edx
  loc_00612439: push eax
  loc_0061243A: call edi
  loc_0061243C: mov ecx, [00668108h]
  loc_00612442: mov edx, [0066810Ch]
  loc_00612448: sub esp, 00000010h
  loc_0061244B: mov eax, esp
  loc_0061244D: push 0043B85Ch ; "Caption"
  loc_00612452: mov [eax], ecx
  loc_00612454: mov ecx, [00668110h]
  loc_0061245A: mov [eax+00000004h], edx
  loc_0061245D: mov edx, [00668114h]
  loc_00612463: mov [eax+00000008h], ecx
  loc_00612466: mov [eax+0000000Ch], edx
  loc_00612469: mov eax, var_30
  loc_0061246C: push eax
  loc_0061246D: call [00401094h] ; __vbaLateMemSt
  loc_00612473: lea ecx, var_30
  loc_00612476: lea edx, var_2C
  loc_00612479: push ecx
  loc_0061247A: lea eax, var_28
  loc_0061247D: push edx
  loc_0061247E: lea ecx, var_24
  loc_00612481: push eax
  loc_00612482: push ecx
  loc_00612483: push 00000004h
  loc_00612485: call [0040104Ch] ; __vbaFreeObjList
  loc_0061248B: mov eax, var_C4
  loc_00612491: add esp, 00000014h
  loc_00612494: mov edx, [eax]
  loc_00612496: push 0000000Ah
  loc_00612498: push eax
  loc_00612499: call [edx+00000048h]
  loc_0061249C: test eax, eax
  loc_0061249E: fnclex
  loc_006124A0: jge 006124B3h
  loc_006124A2: mov ecx, var_C4
  loc_006124A8: push 00000048h
  loc_006124AA: push 0043B5D0h
  loc_006124AF: push ecx
  loc_006124B0: push eax
  loc_006124B1: call edi
  loc_006124B3: mov eax, var_C4
  loc_006124B9: push 0000000Ah
  loc_006124BB: push eax
  loc_006124BC: mov edx, [eax]
  loc_006124BE: call [edx+00000058h]
  loc_006124C1: test eax, eax
  loc_006124C3: fnclex
  loc_006124C5: jge 006124D8h
  loc_006124C7: mov ecx, var_C4
  loc_006124CD: push 00000058h
  loc_006124CF: push 0043B5D0h
  loc_006124D4: push ecx
  loc_006124D5: push eax
  loc_006124D6: call edi
  loc_006124D8: sub esp, 00000010h
  loc_006124DB: mov eax, 00000002h
  loc_006124E0: mov edx, esp
  loc_006124E2: mov ecx, eax
  loc_006124E4: mov var_70, ecx
  loc_006124E7: mov var_68, eax
  loc_006124EA: mov [edx], ecx
  loc_006124EC: mov ecx, var_6C
  loc_006124EF: push 8001000Eh
  loc_006124F4: mov [edx+00000004h], ecx
  loc_006124F7: mov ecx, var_C4
  loc_006124FD: push ecx
  loc_006124FE: mov [edx+00000008h], eax
  loc_00612501: mov eax, var_64
  loc_00612504: mov [edx+0000000Ch], eax
  loc_00612507: call [00401280h] ; __vbaLateIdSt
  loc_0061250D: mov edx, var_C4
  loc_00612513: push 00000000h
  loc_00612515: push 80011003h
  loc_0061251A: push edx
  loc_0061251B: call [0040102Ch] ; __vbaLateIdCall
  loc_00612521: add esp, 0000000Ch
  loc_00612524: lea eax, var_C4
  loc_0061252A: push 00000000h
  loc_0061252C: push eax
  loc_0061252D: call [004010C8h] ; __vbaObjSetAddref
  loc_00612533: push 0061259Dh
  loc_00612538: jmp 00612573h
  loc_0061253A: lea ecx, var_20
  loc_0061253D: call [0040129Ch] ; __vbaFreeStr
  loc_00612543: lea ecx, var_30
  loc_00612546: lea edx, var_2C
  loc_00612549: push ecx
  loc_0061254A: lea eax, var_28
  loc_0061254D: push edx
  loc_0061254E: lea ecx, var_24
  loc_00612551: push eax
  loc_00612552: push ecx
  loc_00612553: push 00000004h
  loc_00612555: call [0040104Ch] ; __vbaFreeObjList
  loc_0061255B: lea edx, var_60
  loc_0061255E: lea eax, var_50
  loc_00612561: push edx
  loc_00612562: lea ecx, var_40
  loc_00612565: push eax
  loc_00612566: push ecx
  loc_00612567: push 00000003h
  loc_00612569: call [0040103Ch] ; __vbaFreeVarList
  loc_0061256F: add esp, 00000024h
  loc_00612572: ret
  loc_00612573: lea edx, var_C8
  loc_00612579: lea eax, var_C4
  loc_0061257F: push edx
  loc_00612580: push eax
  loc_00612581: push 00000002h
  loc_00612583: call [0040104Ch] ; __vbaFreeObjList
  loc_00612589: mov esi, [0040129Ch] ; __vbaFreeStr
  loc_0061258F: add esp, 0000000Ch
  loc_00612592: lea ecx, var_18
  loc_00612595: call __vbaFreeStr
  loc_00612597: lea ecx, var_1C
  loc_0061259A: call __vbaFreeStr
  loc_0061259C: ret
  loc_0061259D: mov ecx, var_10
  loc_006125A0: pop edi
  loc_006125A1: pop esi
  loc_006125A2: mov fs:[00000000h], ecx
  loc_006125A9: pop ebx
  loc_006125AA: mov esp, ebp
  loc_006125AC: pop ebp
  loc_006125AD: ret
End Sub
