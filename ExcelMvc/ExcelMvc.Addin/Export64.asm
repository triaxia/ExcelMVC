EXTERN ExportTable:PTR QWORD
.code

udf0 PROC EXPORT 
   jmp ExportTable + 0 * 8 
udf0 ENDP

end