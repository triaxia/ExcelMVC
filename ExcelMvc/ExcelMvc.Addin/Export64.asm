EXTERN ExportTable:PTR QWORD
.code

f0 PROC EXPORT 
   jmp ExportTable + 0 * 8 
f0 ENDP
f1 PROC EXPORT
   jmp ExportTable + 1 * 8
f1 ENDP
f2 PROC EXPORT
   jmp ExportTable + 2 * 8
f2 ENDP
f3 PROC EXPORT
   jmp ExportTable + 3 * 8
f3 ENDP
f4 PROC EXPORT
   jmp ExportTable + 4 * 8
f4 ENDP
f5 PROC EXPORT
   jmp ExportTable + 5 * 8
f5 ENDP
f6 PROC EXPORT
   jmp ExportTable + 6 * 8
f6 ENDP
f7 PROC EXPORT
   jmp ExportTable + 7 * 8
f7 ENDP
f8 PROC EXPORT
   jmp ExportTable + 8 * 8
f8 ENDP

end