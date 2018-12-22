
Private Sub Proc_13_0_592470(arg_C, arg_10) '592470
  loc_00592470: push ebp
  loc_00592471: mov ebp, esp
  loc_00592473: sub esp, 00000008h
  loc_00592476: push 00405696h ; __vbaExceptHandler
  loc_0059247B: mov eax, fs:[00000000h]
  loc_00592481: push eax
  loc_00592482: mov fs:[00000000h], esp
  loc_00592489: sub esp, 0000004Ch
  loc_0059248C: push ebx
  loc_0059248D: push esi
  loc_0059248E: push edi
  loc_0059248F: mov var_8, esp
  loc_00592492: mov var_4, 00402710h
  loc_00592499: mov esi, Me
  loc_0059249C: mov edi, [0040114Ch] ; __vbaLateIdCallLd
  loc_005924A2: xor eax, eax
  loc_005924A4: lea ecx, var_34
  loc_005924A7: mov var_14, eax
  loc_005924AA: mov var_18, eax
  loc_005924AD: mov var_1C, eax
  loc_005924B0: mov var_20, eax
  loc_005924B3: mov var_24, eax
  loc_005924B6: mov var_34, eax
  loc_005924B9: mov var_44, eax
  loc_005924BC: push eax
  loc_005924BD: mov eax, [esi]
  loc_005924BF: push 00000005h
  loc_005924C1: push eax
  loc_005924C2: push ecx
  loc_005924C3: call edi
  loc_005924C5: mov ebx, [0040121Ch] ; __vbaI4Var
  loc_005924CB: add esp, 00000010h
  loc_005924CE: push eax
  loc_005924CF: call ebx
  loc_005924D1: mov ecx, arg_C
  loc_005924D4: sub eax, 00000001h
  loc_005924D7: jo 0059279Ah
  loc_005924DD: xor edx, edx
  loc_005924DF: cmp ecx, eax
  loc_005924E1: setg dl
  loc_005924E4: neg edx
  loc_005924E6: lea ecx, var_34
  loc_005924E9: mov var_58, edx
  loc_005924EC: call [00401020h] ; __vbaFreeVar
  loc_005924F2: cmp var_58, 0000h
  loc_005924F7: jnz 00592775h
  loc_005924FD: mov eax, [esi]
  loc_005924FF: push 00000000h
  loc_00592501: push 0000000Bh
  loc_00592503: lea ecx, var_34
  loc_00592506: push eax
  loc_00592507: push ecx
  loc_00592508: call edi
  loc_0059250A: add esp, 00000010h
  loc_0059250D: push eax
  loc_0059250E: call ebx
  loc_00592510: lea ecx, var_34
  loc_00592513: mov var_14, eax
  loc_00592516: call [00401020h] ; __vbaFreeVar
  loc_0059251C: mov edx, [esi]
  loc_0059251E: push 00000000h
  loc_00592520: push 0000000Ah
  loc_00592522: lea eax, var_34
  loc_00592525: push edx
  loc_00592526: push eax
  loc_00592527: call edi
  loc_00592529: add esp, 00000010h
  loc_0059252C: push eax
  loc_0059252D: call ebx
  loc_0059252F: lea ecx, var_34
  loc_00592532: mov var_18, eax
  loc_00592535: call [00401020h] ; __vbaFreeVar
  loc_0059253B: mov ecx, [esi]
  loc_0059253D: push 00000000h
  loc_0059253F: push 0000000Dh
  loc_00592541: lea edx, var_34
  loc_00592544: push ecx
  loc_00592545: push edx
  loc_00592546: call edi
  loc_00592548: add esp, 00000010h
  loc_0059254B: push eax
  loc_0059254C: call ebx
  loc_0059254E: lea ecx, var_34
  loc_00592551: mov var_20, eax
  loc_00592554: call [00401020h] ; __vbaFreeVar
  loc_0059255A: mov eax, [esi]
  loc_0059255C: push 00000000h
  loc_0059255E: push 0000000Ch
  loc_00592560: lea ecx, var_34
  loc_00592563: push eax
  loc_00592564: push ecx
  loc_00592565: call edi
  loc_00592567: add esp, 00000010h
  loc_0059256A: push eax
  loc_0059256B: call ebx
  loc_0059256D: lea ecx, var_34
  loc_00592570: mov var_24, eax
  loc_00592573: call [00401020h] ; __vbaFreeVar
  loc_00592579: mov edx, [esi]
  loc_0059257B: push 00000000h
  loc_0059257D: push FFFFFE01h
  loc_00592582: lea eax, var_34
  loc_00592585: push edx
  loc_00592586: push eax
  loc_00592587: call edi
  loc_00592589: add esp, 00000010h
  loc_0059258C: push eax
  loc_0059258D: call ebx
  loc_0059258F: lea ecx, var_34
  loc_00592592: mov var_1C, eax
  loc_00592595: call [00401020h] ; __vbaFreeVar
  loc_0059259B: mov edi, var_40
  loc_0059259E: sub esp, 00000010h
  loc_005925A1: mov edx, esp
  loc_005925A3: mov ecx, 00004003h
  loc_005925A8: mov ebx, var_38
  loc_005925AB: lea eax, arg_C
  loc_005925AE: mov [edx], ecx
  loc_005925B0: push 0000000Bh
  loc_005925B2: mov [edx+00000004h], edi
  loc_005925B5: mov [edx+00000008h], eax
  loc_005925B8: mov eax, [esi]
  loc_005925BA: push eax
  loc_005925BB: mov [edx+0000000Ch], ebx
  loc_005925BE: call [00401280h] ; __vbaLateIdSt
  loc_005925C4: push 00000000h
  loc_005925C6: push 00000006h
  loc_005925C8: mov ecx, [esi]
  loc_005925CA: lea edx, var_34
  loc_005925CD: push ecx
  loc_005925CE: push edx
  loc_005925CF: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005925D5: add esp, 00000010h
  loc_005925D8: push eax
  loc_005925D9: call [0040121Ch] ; __vbaI4Var
  loc_005925DF: sub esp, 00000010h
  loc_005925E2: mov ecx, 00000003h
  loc_005925E7: mov edx, esp
  loc_005925E9: push 0000000Ah
  loc_005925EB: mov [edx], ecx
  loc_005925ED: mov [edx+00000004h], edi
  loc_005925F0: mov [edx+00000008h], eax
  loc_005925F3: mov eax, [esi]
  loc_005925F5: push eax
  loc_005925F6: mov [edx+0000000Ch], ebx
  loc_005925F9: call [00401280h] ; __vbaLateIdSt
  loc_005925FF: lea ecx, var_34
  loc_00592602: call [00401020h] ; __vbaFreeVar
  loc_00592608: sub esp, 00000010h
  loc_0059260B: mov ecx, 00004003h
  loc_00592610: mov edx, esp
  loc_00592612: lea eax, arg_C
  loc_00592615: push 0000000Dh
  loc_00592617: mov [edx], ecx
  loc_00592619: mov [edx+00000004h], edi
  loc_0059261C: mov [edx+00000008h], eax
  loc_0059261F: mov eax, [esi]
  loc_00592621: push eax
  loc_00592622: mov [edx+0000000Ch], ebx
  loc_00592625: call [00401280h] ; __vbaLateIdSt
  loc_0059262B: mov ecx, [esi]
  loc_0059262D: push 00000000h
  loc_0059262F: push 00000004h
  loc_00592631: lea edx, var_34
  loc_00592634: push ecx
  loc_00592635: push edx
  loc_00592636: call [0040114Ch] ; __vbaLateIdCallLd
  loc_0059263C: add esp, 00000010h
  loc_0059263F: push eax
  loc_00592640: call [0040121Ch] ; __vbaI4Var
  loc_00592646: sub eax, 00000001h
  loc_00592649: mov ecx, 00000003h
  loc_0059264E: jo 0059279Ah
  loc_00592654: sub esp, 00000010h
  loc_00592657: mov edx, esp
  loc_00592659: push 0000000Ch
  loc_0059265B: mov [edx], ecx
  loc_0059265D: mov [edx+00000004h], edi
  loc_00592660: mov [edx+00000008h], eax
  loc_00592663: mov eax, [esi]
  loc_00592665: push eax
  loc_00592666: mov [edx+0000000Ch], ebx
  loc_00592669: call [00401280h] ; __vbaLateIdSt
  loc_0059266F: lea ecx, var_34
  loc_00592672: call [00401020h] ; __vbaFreeVar
  loc_00592678: sub esp, 00000010h
  loc_0059267B: mov ecx, 00000003h
  loc_00592680: mov edx, esp
  loc_00592682: mov eax, 00000001h
  loc_00592687: push FFFFFE01h
  loc_0059268C: mov [edx], ecx
  loc_0059268E: mov [edx+00000004h], edi
  loc_00592691: mov [edx+00000008h], eax
  loc_00592694: mov eax, [esi]
  loc_00592696: push eax
  loc_00592697: mov [edx+0000000Ch], ebx
  loc_0059269A: call [00401280h] ; __vbaLateIdSt
  loc_005926A0: sub esp, 00000010h
  loc_005926A3: mov ecx, 00004003h
  loc_005926A8: mov edx, esp
  loc_005926AA: lea eax, arg_10
  loc_005926AD: push 00000026h
  loc_005926AF: mov [edx], ecx
  loc_005926B1: mov [edx+00000004h], edi
  loc_005926B4: mov [edx+00000008h], eax
  loc_005926B7: mov eax, [esi]
  loc_005926B9: push eax
  loc_005926BA: mov [edx+0000000Ch], ebx
  loc_005926BD: call [00401280h] ; __vbaLateIdSt
  loc_005926C3: sub esp, 00000010h
  loc_005926C6: mov ecx, 00004003h
  loc_005926CB: mov edx, esp
  loc_005926CD: lea eax, var_14
  loc_005926D0: push 0000000Bh
  loc_005926D2: mov [edx], ecx
  loc_005926D4: mov [edx+00000004h], edi
  loc_005926D7: mov [edx+00000008h], eax
  loc_005926DA: mov eax, [esi]
  loc_005926DC: push eax
  loc_005926DD: mov [edx+0000000Ch], ebx
  loc_005926E0: call [00401280h] ; __vbaLateIdSt
  loc_005926E6: sub esp, 00000010h
  loc_005926E9: mov ecx, 00004003h
  loc_005926EE: mov edx, esp
  loc_005926F0: lea eax, var_18
  loc_005926F3: push 0000000Ah
  loc_005926F5: mov [edx], ecx
  loc_005926F7: mov [edx+00000004h], edi
  loc_005926FA: mov [edx+00000008h], eax
  loc_005926FD: mov eax, [esi]
  loc_005926FF: push eax
  loc_00592700: mov [edx+0000000Ch], ebx
  loc_00592703: call [00401280h] ; __vbaLateIdSt
  loc_00592709: sub esp, 00000010h
  loc_0059270C: mov ecx, 00004003h
  loc_00592711: mov edx, esp
  loc_00592713: lea eax, var_20
  loc_00592716: push 0000000Dh
  loc_00592718: mov [edx], ecx
  loc_0059271A: mov [edx+00000004h], edi
  loc_0059271D: mov [edx+00000008h], eax
  loc_00592720: mov eax, [esi]
  loc_00592722: push eax
  loc_00592723: mov [edx+0000000Ch], ebx
  loc_00592726: call [00401280h] ; __vbaLateIdSt
  loc_0059272C: sub esp, 00000010h
  loc_0059272F: mov ecx, 00004003h
  loc_00592734: mov edx, esp
  loc_00592736: lea eax, var_24
  loc_00592739: push 0000000Ch
  loc_0059273B: mov [edx], ecx
  loc_0059273D: mov [edx+00000004h], edi
  loc_00592740: mov [edx+00000008h], eax
  loc_00592743: mov eax, [esi]
  loc_00592745: push eax
  loc_00592746: mov [edx+0000000Ch], ebx
  loc_00592749: call [00401280h] ; __vbaLateIdSt
  loc_0059274F: sub esp, 00000010h
  loc_00592752: mov ecx, 00004003h
  loc_00592757: mov edx, esp
  loc_00592759: lea eax, var_1C
  loc_0059275C: push FFFFFE01h
  loc_00592761: mov [edx], ecx
  loc_00592763: mov [edx+00000004h], edi
  loc_00592766: mov [edx+00000008h], eax
  loc_00592769: mov eax, [esi]
  loc_0059276B: push eax
  loc_0059276C: mov [edx+0000000Ch], ebx
  loc_0059276F: call [00401280h] ; __vbaLateIdSt
  loc_00592775: push 00592787h
  loc_0059277A: jmp 00592786h
  loc_0059277C: lea ecx, var_34
  loc_0059277F: call [00401020h] ; __vbaFreeVar
  loc_00592785: ret
  loc_00592786: ret
  loc_00592787: mov ecx, var_10
  loc_0059278A: pop edi
  loc_0059278B: pop esi
  loc_0059278C: mov fs:[00000000h], ecx
  loc_00592793: pop ebx
  loc_00592794: mov esp, ebp
  loc_00592796: pop ebp
  loc_00592797: retn 000Ch
End Sub
