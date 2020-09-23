; Texture.asm by Robert Rayment  June 2001  V.04

; VB
;
; ## ASSEMBLER STRUCTURE FOR TEXTURE.BIN ###################
;
; Public Type PERLININFO  ' Offsets
;   PicWidth As Long     ' 0
;   PicHeight As Long    ' 4
;   ptLS As Long         ' 8   Pointer to LongSurf()
;   NoiseSeed  As Long   ' 12
;   ptRndGrid As Long    ' 16  Pointer to RndGrid()
;   StartScale As Long   ' 20  Perlin noise scales
;   LastScale As Long    ' 24
;   ptPerGrid As Long    ' 28  Pointer to PerGrid()
;   avmax As Long        ' 32  Max/Min values in PerGrid()
;   avmin As Long        ' 36
;   ptzColorWt As Long   ' 40  Pointer to zColorWt()
;   ptzSineAdds As Long  ' 44  Pointer to zSineAdds()
;   WoodType As Long     ' 48
;   MarbleType As Long   ' 52
;	CosLin As Long       ' 56  0 Cosine, 1 Linear
;	CycleSpeed As Long   ' 60  1,2,4
; End Type
; Public Per As PERLININFO
;
; Machine code opcodes
;
; OpCode& = 0 Random noise for RndGrid()
; OpCode& = 1 Perlin noise
; OpCode& = 2 Wood
; OpCode& = 4 Marble
; OpCode& = 8 Smooth
; OpCode& = 16 Cycling
;
; Machine code globals
;
; Global InCode() As Byte    ' For holding machine code
; Global ptMCode             ' Pointer to InCode(1)
; Global ptPerStruc          ' Pointer to Per.PicWidth
;
; Use:-
;
; res& = CallWindowProc(ptMCode, ptPerStruc, ptzres, 2&, OpCode&)
;                                8           12      16  20

;-------------------------------------------------------------
%macro movab 2		;name & num of parameters
  push dword %2		;2nd param
  pop dword %1		;1st param
%endmacro			;use  movab %1,%2
;Allows eg	movab bmW,[ebx+4]
;-------------------------------------------------------------

; INPUT DATA STORE
; Long integers
%define PicWidth	[ebp-4]
%define PicHeight	[ebp-8]
%define ptLS        [ebp-12]
%define NoiseSeed   [ebp-16]
%define ptRndGrid   [ebp-20]
%define StartScale  [ebp-24]
%define LastScale   [ebp-28]
%define ptPerGrid   [ebp-32]
%define avmax       [ebp-36]
%define avmin       [ebp-40]
%define ptzColorWt  [ebp-44]
%define ptzSineAdds [ebp-48]
%define WoodType    [ebp-52]
%define MarbleType  [ebp-56]   
%define CosLin		[ebp-60]   
%define CycleSpeed	[ebp-64]   



[bits 32]

	push ebp
	mov ebp,esp
	sub esp,196
	push edi
	push esi
	push ebx

;	LOAD INPUT DATA

	mov ebx,[ebp+8]		; ->ptPerStruc
	
	movab PicWidth,		[ebx]
	movab PicHeight,	[ebx+4]
	movab ptLS,			[ebx+8]
	movab NoiseSeed,	[ebx+12]
	movab ptRndGrid,	[ebx+16]
	movab StartScale,	[ebx+20]
	movab LastScale,	[ebx+24]
	movab ptPerGrid,	[ebx+28]
	movab avmax,		[ebx+32]
	movab avmin,		[ebx+36]
	movab ptzColorWt,	[ebx+40]
	movab ptzSineAdds,	[ebx+44]
	movab WoodType,		[ebx+48]
	movab MarbleType,	[ebx+52]
	movab CosLin,		[ebx+56]
	movab CycleSpeed,	[ebx+60]

; ----------------------------
;
; OpCode& = 0  Random noise for RndGrid()
; OpCode& = 1  Perlin noise
; OpCode& = 2  Wood
; OpCode& = 4  Marble
; OpCode& = 8  Smooth
; OpCode& = 16 Cycle colors
	
; GET OpCode&

	mov eax,[ebp+20]
	cmp eax,0
	jne TestFor1
	CALL RANDNOISE
	jmp GETOUT
TestFor1:
	rcr eax,1
	jnc TestFor2
	CALL RANDNOISE
	CALL PERLIN
	jmp GETOUT
TestFor2:
	rcr eax,1
	jnc TestFor4
	CALL SINENOISE
	CALL WOOD
	jmp GETOUT
TestFor4:
	rcr eax,1
	jnc TestFor8
	CALL RANDNOISE
	CALL PERLIN
	CALL MARBLE
	jmp GETOUT
TestFor8:
	rcr eax,1
	jnc TestFor16
	CALL SMOOTH
	jmp GETOUT

TestFor16:
	rcr eax,1
	jnc TestFor32
	CALL CYCLE
	jmp GETOUT

TestFor32:

GETOUT:


	pop ebx
	pop esi
	pop edi
	mov esp,ebp
	pop ebp
	RET 16

;-----------------------------------------------------------------
	; WAIT FOR SCREEN REFRESH TO FINISH
	; Need 3 waits to be sure of getting
	; maximum time for re-drawing screen
	; Probably not necessary when using
	; StretchDIBits
	
	mov   dx,03DAh		; Status register
	
WaitForOFF:
	in    al,dx			; Will NT allow this 'in' instruction ??
	test  al,8			; Is scan OFF
	jnz   WaitForOFF	; No, keep waiting
	
WaitForON:
	in    al,dx     
	test  al,8	  		; Is scan ON
	jz    WaitForON		; No, keep waiting
	
WaitForOFF2:
	in    al,dx     
	test  al,8			; Is scan OFF
	jnz   WaitForOFF2	; No, keep waiting
;-----------------------------------------------------------------


	ret 16
;=====================================================

RANDNOISE:

	mov edi,ptRndGrid	; ptr to RndGrid(1,1)
	mov eax,PicWidth
	mov ebx,PicHeight
	mul ebx
	mov ecx,eax
nexran:
	mov eax,011813h	 	; 71699 prime 
	imul DWORD NoiseSeed
	add eax, 0AB209h 	; 700937 prime
	rcr eax,1			; leaving out gives vertical lines plus
						; faint horizontal ones, tartan

	;----------------------------------------
	;jc ok				; these 2 have little effect
	;rol eax,1			;
ok:						;
	;----------------------------------------
	
	;----------------------------------------
	;dec eax			; these produce vert lines
	;inc eax			; & with fsin marble arches
	;----------------------------------------

	mov NoiseSeed,eax	; save seed
	
	;----------------------------------------
	;fild dword NoiseSeed	; makes diagonals, repetitive marble
	;fsin					; fcos smooth no structure
	;fistp dword NoiseSeed
	;----------------------------------------
	
	;and eax,0FFh		; can leave out but Fh 7h 3h 2h 1h all work
						; giving similar but different fuzzy patterns
						; Suggests that all that may be required is a
						; random pattern of 1's and 0's.
						
	stosb				; mov [edi],AL, inc edi 

	dec ecx
	jnz nexran
	
RET
;=====================================================

PERLIN:

%define sca		[ebp-68]	; sca init = StartScale
%define j		[ebp-72]	; j = PicHeight To 2 Step -2
%define jg		[ebp-76]	; jg = ((j-1) \ sca) + 1   IntDiv
%define i		[ebp-80]	; i = PicWidth To 2 Step -2
%define ig		[ebp-84]	; ig = ((i-1) \ sca) + 1   IntDiv
%define ix0		[ebp-88]	; ix0 = (ig-1) * sca
%define iy0		[ebp-92]	; iy0 = (jg-1) * sca
%define ix1		[ebp-96]	; ix1 = ig * sca
%define iy1		[ebp-100]	; iy1 = jg * sca

%define X0		[ebp-104]
%define Y0		[ebp-108]
%define X1		[ebp-112]
%define Y1		[ebp-116]

%define zs		[ebp-120]
%define zt		[ebp-124]
%define zu		[ebp-128]
%define zv		[ebp-132]

%define zSx		[ebp-136]
%define zSy		[ebp-140]

%define PerGridij	[ebp-144]		; PerGrid(i,j)
%define xavyavsca	[ebp-148]		; (xav + yav) * sca

%define zmult		[ebp-152]		; 256/(avmax - avmin)
%define cul			[ebp-156]

%define num2		[ebp-160]		; 2
%define num3		[ebp-164]		; 3
%define num255		[ebp-168]		; 255

%define zColorWt0	[ebp-172]
%define zColorWt1	[ebp-176]
%define zColorWt2	[ebp-180]

%define iculr		[ebp-184]
%define iculg		[ebp-188]
%define iculb		[ebp-192]
%define zp			[ebp-196]		; pi#/2


	mov eax,100000
	mov avmin,eax
	neg eax
	mov avmax,eax

	mov eax,2
	mov num2,eax
	inc eax
	mov num3,eax
	mov eax,255
	mov num255,eax
	
	fldpi
	fild dword num2
	fdivp st1			; st1/st0 = pi/2
	fstp dword zp

	mov eax,StartScale

MAINDO:
	mov sca,eax
	mov eax,PicHeight
FORJ:
	mov j,eax
	dec eax				; (j-1)
	mov ebx,sca
	mov edx,0
	div ebx				; (j-1)\sca  ;;OVF
	inc eax
	mov jg,eax			; jg = ((j-1)\sca) + 1

	mov eax,PicWidth
FORI:
	mov i,eax
	dec eax				; (i-1)
	mov ebx,sca
	mov edx,0
	div ebx				; (i-1)\sca
	inc eax
	mov ig,eax			; ig = ((i-1)\sca) + 1
		
	mov eax,ig			; Get pix coords from RndGrid coords
	dec eax
	mov ebx,sca
	mul ebx
	mov ix0,eax			; ix0 = (ig-1)*sca

	mov eax,jg
	dec eax
	mov ebx,sca
	mul ebx
	mov iy0,eax			; iy0 = (jg-1)*sca

	mov eax,ig
	mov ebx,sca
	mul ebx
	mov ix1,eax			; ix1 = ig*sca
	
	mov eax,jg
	mov ebx,sca
	mul ebx
	mov iy1,eax			; iy1 = jg*sca

	mov eax,CosLin
	cmp eax,1
	JE Near LINEAR		; Linear Else Cosine smoothing
	
	; COSINE smoothing -------------
	fild dword i		
	fild dword ix0
	fsubp st1			; st1-st0 = i-ix0
	fld dword zp
	fmulp st1			; zp*(i-ix0)
	fild dword sca
	fdivp st1			; st1/st0 = zp*(i-ix0)/sca
	fcos
	fstp dword X0		; X0 = Cos(zp*(i-ix0)/sca)

	fild dword j		
	fild dword iy0
	fsubp st1			; st1-st0 = j-iy0
	fld dword zp
	fmulp st1			; zp*(j-iy0)
	fild dword sca
	fdivp st1			; st1/st0 = zp*(j-iy0)/sca
	fcos
	fstp dword Y0		; Y0 = Cos(zp*(j-iy0)/sca)
	
	fild dword ix1
	fild dword i
	fsubp st1			; st1-st0 = ix1-i
	fld dword zp
	fmulp st1			; zp*(ix1-i)
	fild dword sca
	fdivp st1			; st1/st0 = zp*(ix1-i)/sca
	fcos
	fstp dword X1		; X1 = Cos(zp*(ix1-i)/sca)
	
	fild dword iy1
	fild dword j
	fsubp st1			; st1-st0 = iy1-j
	fld dword zp
	fmulp st1			; zp*(iy1-j)
	fild dword sca
	fdivp st1			; st1/st0 = zp*(iy1-j)/sca
	fcos
	fstp dword Y1		; Y1 = Cos(zp*(iy1-j)/sca)
	; End Cosine smoothing -------------

	; Cosine Vector dot, RndGrid() bytes -------
	mov eax,jg
	dec eax
	mov ebx,PicWidth
	mul ebx
	mov ebx,ig
	dec ebx
	add eax,ebx
	mov edi,ptRndGrid
	add edi,eax			; -> RndGrid(ig,jg)
	movzx eax,byte[edi]
	mov zs,eax			; RndGrid(ig,jg)
	fild dword zs
	fld dword X0
	fld dword Y0
	faddp st1
	fld1
	fsubp st1			; st1-st0 = (X0+Y0-1)
	fmulp st1
	fstp dword zs		; zs = RndGrid(ig,jg)*(X0+Y0-1)
	
	inc edi
	movzx eax,byte[edi]
	mov zt,eax			; RndGrid(ig+1,jg)
	fild dword zt
	fld dword X1
	fld dword Y0
	faddp st1
	fld1
	fsubp st1			; st1-st0 = (X1+Y0-1)
	fmulp st1
	fstp dword zt		; zt = RndGrid(ig+1,jg)*(X1+Y0-1)

	dec edi				; -> RndGrid(ig,jg)
	mov eax,PicWidth
	add edi,eax			; -> RndGrid(ig,jg+1)
	movzx eax,byte[edi]
	mov zu,eax			; RndGrid(ig,jg+1)
	fild dword zu
	fld dword X0
	fld dword Y1
	faddp st1
	fld1
	fsubp st1			; st1-st0 = (X0+Y1-1)
	fmulp st1
	fstp dword zu		; zu = RndGrid(ig,jg+1)*(X0+Y1-1)
	
	inc edi				; -> RndGrid(ig+1,jg+1)
	movzx eax,byte[edi]
	mov zv,eax
	fild dword zv
	fld dword X1
	fld dword Y1
	faddp st1
	fld1
	fsubp st1			; st1-st0 = (X1+Y1-1)
	fmulp st1
	fstp dword zv		; zv = RndGrid(ig+1,jg+1)*(X1+Y1-1)
	; End Vector dot -----------------------

	; EASE curve ---------------------------
	fild dword i		; Restore X0
	fild dword ix0
	fsubp st1			; st1-st0 = i-ix0
	fild dword sca
	fdivp st1			; st1/st0 = (i-ix0)/sca
	fstp dword X0		; X0 = (i-ix0)/sca
	fild dword j		; Restore Y0
	fild dword iy0
	fsubp st1			; st1-st0 = j-iy0
	fild dword sca
	fdivp st1			; st1/st0 = (j-iy0)/sca
	fstp dword Y0		; Y0 = (j-iy0)/sca
	;----------------------
	fild dword num3		; zSx = 3*X0^2 - 2*X0^3
	fld dword X0
	fld dword X0
	fmulp st1			; X0^2
	fmulp st1			; 3*X0^2
	fild dword num2
	fld dword X0
	fld dword X0
	fmulp st1			; X0^2
	fld dword X0
	fmulp st1			; X0^3
	fmulp st1			; 2*X0^3
	fsubp st1			; st1-st0 = zSx = 3*X0^2 - 2*X0^3
	fst dword zSx		; zSx,zSx
	
	fld dword zt
	fld dword zs
	fsubp st1			; st1-st0 = zt-zs
	fmulp st1			; zSx*(zt-zs)
	fld dword zs
	faddp st1			; xa = zs + zSx*(zt-zs)
	
	fld dword zSx
	fld dword zv
	fld dword zu
	fsubp st1			; st1-st0 = zv-zu
	fmulp st1			; zSx*(zv-zu)
	fld dword zu
	faddp st1			; xb = zu + zSx*(zv-zu)
	
	faddp st1			; xa + xb
	fild dword num2
	fdivp st1			; st1/sto = (xa+xb)/2
	
	fild dword num3		; zSy = 3*Y0^2 - 2*Y0^3
	fld dword Y0
	fld dword Y0
	fmulp st1			; Y0^2
	fmulp st1			; 3*Y0^2
	fild dword num2
	fld dword Y0
	fld dword Y0
	fmulp st1			; Y0^2
	fld dword Y0
	fmulp st1			; Y0^3
	fmulp st1			; 2*Y0^3
	fsubp st1			; st1-st0 = zSy = 3*Y0^2 - 2*Y0^3
	fst dword zSy		; zSy,zSy
		
	fld dword zu
	fld dword zs
	fsubp st1			; st1-st0 = zu-zs
	fmulp st1			; zSy*(zu-zs)
	fld dword zs
	faddp st1			; ya = zs + zSy*(zu-zs)
		
	fld dword zSy
	fld dword zv
	fld dword zt
	fsubp st1			; st1-st0 = zv-zt
	fmulp st1			; zSy*(zv-zt)
	fld dword zt
	faddp st1			; xb = zt + zSy*(zv-zt)
		
	faddp st1			; ya + yb
	fild dword num2
	fdivp st1			; st1/sto = (ya+yb)/2
	
	faddp st1			; (xa+xb)/2 + (ya+yb)/2
	fild dword sca
	fmulp st1			; ((xa+xb)/2 + (ya+yb)/2) * sca
	
	fistp dword xavyavsca	; = ((xa+xb)/2 + (ya+yb)/2) * sca
	; End EASE curve ---------------------------
	JMP AVAV


LINEAR:
	; LINEAR Vector dot, RndGrid() bytes -------

	;   sca2zs = RndGrid(ig, jg) * (sca - i + ix0) * (sca - j + iy0)
	;   sca2zt = RndGrid(ig + 1, jg) * (sca - ix1 + i) * (sca - j + iy0)
	;   sca2zu = RndGrid(ig, jg + 1) * (sca - i + ix0) * (sca - iy1 + j)
	;   sca2zv = RndGrid(ig + 1, jg + 1) * (sca - ix1 + i) * (sca - iy1 + j)
	;      
	;   izav = (sca2zs + sca2zt + sca2zu + sca2zv) / sca   ' in effect * Sca to decrease the
	;                                                      ' amp of higher freqs
	;      
	;  PerGrid(i, j) = PerGrid(i, j) + izav

	mov eax,jg
	dec eax
	mov ebx,PicWidth
	mul ebx
	mov ebx,ig
	dec ebx
	add eax,ebx
	mov edi,ptRndGrid
	add edi,eax			; -> RndGrid(ig,jg)
	movzx eax,byte[edi]
	
;   sca2zs = RndGrid(ig, jg) * (sca - i + ix0) * (sca - j + iy0)
	mov ebx,sca
	sub ebx,i			; sca-i
	add ebx,ix0			; sca-i+ix0
	mul ebx
	mov ebx,sca
	sub ebx,j			; sca-j
	add ebx,iy0			; sca-j+iy0
	mul ebx
	mov zs,eax
	;------------------	
	
	inc edi
	movzx eax,byte[edi]

;   sca2zt = RndGrid(ig + 1, jg) * (sca - ix1 + i) * (sca - j + iy0)
	mov ebx,sca
	sub ebx,ix1
	add ebx,i
	mul ebx
	mov ebx,sca
	sub ebx,j
	add ebx,iy0
	mul ebx
	mov zt,eax
	;------------------	

	dec edi				; -> RndGrid(ig,jg)
	mov eax,PicWidth
	add edi,eax			; -> RndGrid(ig,jg+1)
	movzx eax,byte[edi]

;   sca2zu = RndGrid(ig, jg + 1) * (sca - i + ix0) * (sca - iy1 + j)
	mov ebx,sca
	sub ebx,i
	add ebx,ix0
	mul ebx
	mov ebx,sca
	sub ebx,iy1
	add ebx,j
	mul ebx
	mov zu,eax
	;------------------	

	inc edi				; -> RndGrid(ig+1,jg+1)
	movzx eax,byte[edi]

;   sca2zv = RndGrid(ig + 1, jg + 1) * (sca - ix1 + i) * (sca - iy1 + j)
	mov ebx,sca
	sub ebx,ix1
	add ebx,i
	mul ebx
	mov ebx,sca
	sub ebx,iy1
	add ebx,j
	mul ebx
	mov zv,eax

;   izav = (sca2zs + sca2zt + sca2zu + sca2zv) / sca
;   PerGrid(i, j) = PerGrid(i, j) + izav
	mov eax,zs
	add eax,zt
	add eax,zu
	add eax,zv
	mov ebx,sca
	div ebx
	mov xavyavsca,eax
	
AVAV:		; Finalising Linear & Cosine smoothing
	; Fill PerGrid() Long integer --------------
	mov eax,j
	dec eax
	mov ebx,PicWidth
	mul ebx				; eax = (j-1)*PicWidth
	mov ebx,i
	dec ebx				; ebx = (i-1)
	add eax,ebx			; eax = (j-1)*PicWidth+(i-1) 
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	mov edi,ptPerGrid
	add edi,eax			; -> PerGrid(i,j)

	mov eax,[edi]
	add eax,xavyavsca
	mov [edi],eax		; PerGrid(i,j) = PerGrid(i,j) + xavyavsca
	mov PerGridij,eax	; PerGrid(i,j)
	mov edx,eax			; edx=PerGrid(i,j)
	
	mov eax,4			; Fill in odd points
	sub edi,eax			; -> PerGrid(i-1,j)
	mov [edi],edx
	add edi,eax			; -> PerGrid(i,j)
	mov eax,PicWidth
	shl eax,2			; 4*PicWidth
	sub edi,eax			; -> PerGrid(i,j-1)
	mov [edi],edx
	mov eax,4
	sub edi,eax			; -> PerGrid(i-1,j-1)
	mov [edi],edx
	; End Fill PerGrid() Long integer --------------
	
	; Find max/min ---------------------------------
	mov eax,PerGridij
	cmp eax,avmin		; eax-avmin
	jge Testmax
	mov avmin,eax
	jmp NEXTI
Testmax:
	cmp eax,avmax		; eax-avmax
	jle NEXTI
	mov avmax,eax
	; End Find max/min ---------------------------------
	
NEXTI:
	mov eax,i
	dec eax
	dec eax
	cmp eax,2
	jge Near FORI

	mov eax,j
	dec eax
	dec eax
	cmp eax,2
	jge Near FORJ
	
	; Decrease scale
	mov eax,sca
	shr eax,1			; sca\2
	cmp eax,LastScale
	jge Near MAINDO
	;------------------------------------

	; Scale color -----------------------
	fild dword num255
	fild dword avmax
	fild dword avmin
	fsubp st1			; st1-st0 = avmax-avmin
	fdivp st1			; st1/st0 = 255/(avmax-avmin)
	fstp dword zmult
	; End Scale color -----------------------
	
	; Get zColorWt0,1,2 -----------------------
	mov edi,ptzColorWt
	fld dword [edi]
	fstp dword zColorWt0
	fld dword [edi+4]
	fstp dword zColorWt1
	fld dword [edi+8]
	fstp dword zColorWt2
	; End Get zColorWt0,1,2 -----------------------
	
	; Set pixels from scaled PerGrid(i,j) -------

	; Calc LongSurf(PicWidth,PicHeight) address
	mov eax,PicHeight
	dec eax
	mov ebx,PicWidth	; i
	mul ebx				; eax = (j-1)*PicWidth
	dec ebx				; ebx = (i-1)
	add eax,ebx
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	
	mov edi,ptLS		; -> LongSurf(1,1)
	add edi,eax			; edi -> LongSurf(PicWidth,PicHeight)

	std					; decr edi for stosd
	
	; Calc PerGrid(PicWidth,PicHeight) address
	mov eax,PicHeight
	dec eax
	mov ebx,PicWidth
	mul ebx				; eax = (j-1)*PicWidth
	dec ebx				; (i-1)
	add eax,ebx
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	mov esi,ptPerGrid	; -> PerGrid(1,1)
	add esi,eax			; esi -> PerGrid(PicWidth,PicHeight)

	mov eax,PicHeight
FJ:
	mov j,eax
	mov ecx,PicWidth
FI:
	fild dword [esi]	; PerGrid(i,j)
	fild dword avmin
	fsubp st1			; st1-sto = PerGrid(i,j)-avmin
	fld dword zmult
	fmulp st1			; cul = (PerGrid(i,j)-avmin) * zmult
	fst dword cul		; cul, cul

	fld dword zColorWt0	; Multiply RGB by color weights
	fmulp st1
	fistp dword iculr
	
	fld dword cul
	fld dword zColorWt1
	fmulp st1
	fistp dword iculg
	
	fld dword cul
	fld dword zColorWt2
	fmulp st1
	fistp dword iculb
	
	movzx eax,byte iculb	; Make long color
	movzx ebx,byte iculg
	shl ebx,8
	or eax,ebx
	movzx ebx,byte iculr
	shl ebx,16
	or eax,ebx				; eax = HiB ARGB LoB

	stosd				; mov [edi],eax & edi-4
	
	mov eax,4
	sub esi,eax			; esi-4 Next PerGrid()

	dec ecx
	jnz Near FI
	
	mov eax,j
	dec eax
	jnz Near FJ

	cld					; Restore direction flag

	mov ebx,[ebp+8]		; ->ptPerStruc
	movab [ebx+32],avmax
	movab [ebx+36],avmin

RET
;=====================================================

SINENOISE:

%define T1		[ebp-68]	; T1 = 4 * PicWidth
%define T2		[ebp-72]	; T2 = T1/2
%define T4		[ebp-76]	; T3 = T1/4
%define num2	[ebp-80]	; 2
%define num3	[ebp-84]	; 3
%define IX		[ebp-88]	; IX
%define zp2		[ebp-92]	; 2*pi#

	mov eax,PicWidth
	shl eax,2
	mov T1,eax			; 4*PicWidth
	shr eax,1
	mov T2,eax			; T1/2
	shr eax,1
	mov T4,eax			; T1/4
	
	mov eax,2
	mov num2,eax
	inc eax
	mov num3,eax
	
	fldpi
	fild dword num2
	fmulp st1
	fstp dword zp2		; 2*pi
	
	; Get addresses to zSineAdds()
	mov eax,T2		; T1/2
	dec eax
	shl eax,2		; 4*T2
	mov edi,ptzSineAdds
	add edi,eax		; ->zSineAdds(IX)
	
	mov eax,PicWidth
	shl eax,2
	mov ebx,T2
	sub eax,ebx		; 4*PicWidth-IX
	shl eax,2		; 4*(4*PicWidth-IX)
	mov esi,ptzSineAdds
	add esi,eax		; ->zSineAdds(4*PicWidth+1-IX)
	
	
	mov ecx,T2
FORIX:
	mov IX,ecx
	
	fld dword zp2
	fild dword IX
	fmulp st1
	fild dword T1
	fdivp st1		;st1/sto = zp*IX/T1
	fsin
	
	fld dword zp2
	fild dword IX
	fmulp st1
	fild dword T2
	fdivp st1		;st1/sto = zp*IX/T2
	fsin
	
	fld dword zp2
	fild dword IX
	fmulp st1
	fild dword T4
	fdivp st1		;st1/sto = zp*IX/T4
	fsin
	
	faddp st1		; Y1+Y2
	faddp st1		; Y1+Y2+Y3
	fild dword num3
	fdivp st1		; st1/st0 = (Y1+Y2+Y3)/3
	
	fst dword [edi]
	fstp dword [esi]
	
	mov eax,4
	sub edi,eax
	add esi,eax
	
	dec ecx
	jnz FORIX

RET
;=====================================================

WOOD:

%define AMP		[ebp-68]	; 16
%define ixx		[ebp-72]
%define iyy		[ebp-76]
%define zr		[ebp-80]
%define zalpha	[ebp-84]
%define zSA		[ebp-88]
%define zs		[ebp-92]
%define zRed	[ebp-96]
%define zGreen	[ebp-100]
%define zBlue	[ebp-104]
%define iRed	[ebp-108]
%define iGreen	[ebp-112]
%define iBlue	[ebp-116]
%define IX		[ebp-120]
%define IY		[ebp-124]
%define zColWt0	[ebp-128]
%define zColWt1	[ebp-132]
%define zColWt2	[ebp-136]
%define num16	[ebp-140]	; 16
%define nm2		[ebp-144]	; 2
%define zp17	[ebp-148]	; .17
%define zp08	[ebp-152]	; .08
%define numb100	[ebp-156]	; 100
%define z1mr	[ebp-160]	; (1-zr)
%define nm255	[ebp-164]	; 255
%define nsa		[ebp-168]	; for zSineAdds(nsa)

	mov eax,1
	mov nsa,eax
	
	mov eax,16
	mov AMP,eax
	
	mov eax,16
	mov num16,eax
	
	mov eax,2
	mov nm2,eax
	
	mov eax,100
	mov numb100,eax
	
	; Multipliers for color weightings
	mov eax,17
	mov zp17,eax
	fild dword zp17
	fild dword numb100
	fdivp st1			; st1/sto = 17/100 = .17
	fstp dword zp17
	
	mov eax,8
	mov zp08,eax
	fild dword zp08
	fild dword numb100
	fdivp st1			; st1/sto = 8/100 = .08
	fstp dword zp08
	
	mov eax,255
	mov nm255,eax
	
	
	; Get zColorWt0,1,2 -----------------------
	mov edi,ptzColorWt
	fld dword [edi]
	fstp dword zColWt0
	fld dword [edi+4]
	fstp dword zColWt1
	fld dword [edi+8]
	fstp dword zColWt2
	; End Get zColorWt0,1,2 -----------------------

	; Calc LongSurf(PicWidth,PicHeight) address
	mov eax,PicHeight
	dec eax
	mov ebx,PicWidth	; i
	mul ebx				; eax = (j-1)*PicWidth
	dec ebx				; ebx = (i-1)
	add eax,ebx
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	
	mov edi,ptLS		; -> LongSurf(1,1)
	add edi,eax			; edi -> LongSurf(PicWidth,PicHeight)

	std					; decr edi for stosd
	
	mov eax,PicHeight
FRIY:
	mov IY,eax
	
	mov eax,nsa			; Get zSA
	dec eax
	shl eax,2			; 4*(nsa-1)
	mov esi,ptzSineAdds
	add esi,eax			; ->zSineAdds(nsa)
	fld dword [esi]
	fild dword AMP
	fmulp st1			; zSA
	fstp dword zSA

	mov eax,PicWidth
FRIX:
	mov IX,eax
	
	mov ebx,PicWidth
	shr ebx,1			; ebx = PicWidth/2
	sub eax,ebx
	mov ixx,eax
	
	mov eax,IY
	mov ebx,PicHeight
	shr ebx,1			; ebx = PicHeight/2
	sub eax,ebx
	mov iyy,eax

	mov eax,WoodType
	cmp eax,0
	jne TestWoodType1
	
	fild dword ixx		; WoodType 0 RINGS
	fild dword ixx
	fmulp st1	
	fild dword iyy
	fild dword iyy
	fmulp st1	
	faddp st1
	fsqrt
	fstp dword zr		; zr = Sqr(ixx^2+iyy^2)
	
	jmp GetAlpha
TestWoodType1:
	cmp eax,1
	jne  TestWoodType2

	fild dword ixx		; WoodType 1 LINES
	fild dword iyy
	faddp st1
	fabs
	fstp dword zr		; zr = Abs(ixx+iyy)
	
	jmp GetAlpha
TestWoodType2:

	fild dword ixx		; WoodType 2 CELLS
	fild dword ixx
	fmulp st1	
	fild dword iyy
	fild dword iyy
	fmulp st1	
	faddp st1
	fild dword num16
	fdivp st1			; st1/st0 = Abs(ixx^2+iyy^2)/16
	fstp dword zr		; zr = Abs(ixx^2+iyy^2)/16
	
GetAlpha:
	fild dword iyy
	fild dword ixx
	fpatan				; atan(st1/st0) = atan(iyy/ixx)
	fstp dword zalpha
	
	fld dword zalpha	; zalpha
	fild dword nm2
	fmulp st1
	fsin				; sin(2*zalpha)
	fld dword zr
	fmulp st1			; zr*sin(2*zalpha)
	fld dword zSA
	fdivp st1			; st1/st0 = zs = Sin(2*zalpha) * zr/zSA 
	
	fld dword zr
	faddp st1
	fabs				; zr = Abs(zr + zs) 
	
	fistp dword zr		; izr
	mov eax,zr
	and eax,0Fh			; izr Mod 16
	mov zr,eax
	fild dword zr
	fild dword num16
	fdivp st1			; st1/st0 = zr = (zr Mod 16)/16 (0-1)
	fstp dword zr
	
	fld1
	fld dword zr
	fsubp st1			; st1-st0 = 1-zr
	fst dword z1mr
	
	fld dword zp17
	fmulp st1			; 0.17*(1-zr)
	fld dword zr
	faddp st1			; zr + 0.17*(1-zr)
	fld dword zColWt0
	fmulp st1			; zColorWt(0)*(zr + 0.17*(1-zr))
	fild dword nm255
	fmulp st1
	fistp dword iRed	; zRed*255 
	
	fld dword z1mr		; (1-zr)
	fld dword zp08		
	fmulp st1			; 0.08*(1-zr)
	fld dword zr
	faddp st1			; zr + 0.08*(1-zr)
	fld dword zColWt1
	fmulp st1			; zColorWt(1)*(zr + 0.08*(1-zr))
	fild dword nm255
	fmulp st1
	fistp dword iGreen	; zGreen*255 

	fld dword z1mr		; (1-zr)
	fld dword zp17		
	fmulp st1			; 0.17*(1-zr)
	fld dword zr
	faddp st1			; zr + 0.17*(1-zr)
	fld dword zColWt2
	fmulp st1			; zColorWt(1)*(zr + 0.17*(1-zr))
	fild dword nm255
	fmulp st1
	fistp dword iBlue	; zBlue*255 
	
	; Make up long color
	mov eax,iBlue
	and eax,255
	mov ebx,iGreen
	and ebx,255
	shl ebx,8
	or eax,ebx
	mov ebx,iRed
	and ebx,255
	shl ebx,16
	or eax,ebx			; eax =  HiB ARGB LoB
	
	stosd				; mov [edi],eax & edi-4

	mov eax,IX
	dec eax
	jnz Near FRIX
	
	mov eax,nsa
	inc eax				; nsa+1
	mov ebx,PicWidth
	shl ebx,2			; 4*PicWidth
	cmp eax,ebx			; nsa-4*PicWidth
	jl nsaOK
	mov eax,1
nsaOK:
	mov nsa,eax
	
	mov eax,IY
	dec eax
	jnz Near FRIY
	
	cld					; Restore direction flag
RET
;=====================================================

MARBLE:

%define zAMP		[ebp-68]	; .5
%define zSinVal		[ebp-72]	; zAmp*Sin(zN*pi#+1)
%define MiRed		[ebp-76]
%define MiGreen		[ebp-80]
%define MiBlue		[ebp-84]
%define zMColWt0	[ebp-88]
%define zMColWt1	[ebp-92]
%define zMColWt2	[ebp-96]
%define z2p			[ebp-100]	; 2
%define zp65		[ebp-104]	; .65
%define z2p1		[ebp-108]	; 2.1
%define num100		[ebp-112]	; 100
%define nm255M		[ebp-116]	; 255
%define avmaxmin	[ebp-120]	; avmax-avmin
%define NMUL		[ebp-124]	; MarbleType 0 4, 1 12, 2 20

	mov eax,100
	mov num100,eax
	
	mov eax,255
	mov nm255M,eax

	; zAMP
	mov eax,50
	mov zAMP,eax
	fild dword zAMP
	fild dword num100
	fdivp st1			; st1/sto = 50/100 = .5
	fstp dword zAMP
	
	; Multipliers for color weightings
	mov eax,2
	mov z2p,eax
	fild dword z2p
	fstp dword z2p		; 2.0
	
	mov eax,65
	mov zp65,eax
	fild dword zp65
	fild dword num100
	fdivp st1			; st1/sto = 65/100 = .65
	fstp dword zp65

	mov eax,210
	mov z2p1,eax
	fild dword z2p1
	fild dword num100
	fdivp st1			; st1/sto = 210/100 = 2.1
	fstp dword z2p1
	
	; Get zColorWt0,1,2 -----------------------
	mov edi,ptzColorWt
	fld dword [edi]
	fstp dword zMColWt0
	fld dword [edi+4]
	fstp dword zMColWt1
	fld dword [edi+8]
	fstp dword zMColWt2
	; End Get zColorWt0,1,2 -----------------------

	; Set NMUL according to MarbleType
	mov eax,MarbleType
	cmp eax,0
	jne TestMarble1
	mov eax,4
	jmp SaveMarbleType
TestMarble1:
	cmp eax,1
	jne TestMarble2
	mov eax,12
	jmp SaveMarbleType
TestMarble2:
	mov eax,20
SaveMarbleType:
	mov NMUL,eax
	
	; Calc (avmax-avmin)	
	mov eax,avmax
	mov ebx,avmin
	sub eax,ebx				; avmax-avmin
	mov avmaxmin,eax
	

	; Calc PerGrid(PicWidth,PicHeight) address
	mov eax,PicHeight
	dec eax
	mov ebx,PicWidth
	mul ebx				; eax = (j-1)*PicWidth
	dec ebx				; (i-1)
	add eax,ebx
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	mov esi,ptPerGrid	; -> PerGrid(1,1)
	add esi,eax			; esi -> PerGrid(PicWidth,PicHeight)

	; Calc LongSurf(PicWidth,PicHeight) address
	mov eax,PicHeight
	dec eax
	mov ebx,PicWidth	; i
	mul ebx				; eax = (j-1)*PicWidth
	dec ebx				; ebx = (i-1)
	add eax,ebx
	shl eax,2			; eax = 4 * [(j-1)*PicWidth+(i-1)]
	
	mov edi,ptLS		; -> LongSurf(1,1)
	add edi,eax			; edi -> LongSurf(PicWidth,PicHeight)

	std					; decr edi for stosd
	
	mov edx,PicHeight
MFRIY:
	
	mov ecx,PicWidth
MFRIX:
	
	fild dword [esi]	; PerGrid(IX,IY)
	fild dword avmin
	fsubp st1			; st1-st0 = PerGrid(IX,IY)-avmin
	fild dword avmaxmin
	fdivp st1			; zN = (PerGrid(IX,IY)-avmin)/(avmax-avmin)
	
	fild dword NMUL
	fmulp st1			; zN = NMUL * zN
	fldpi
	fmulp st1			; zN*pi
	fld1
	faddp st1			; zN*pi + 1
	fsin
	fld dword zAMP
	fmulp st1
	fstp dword zSinVal	; zSinVal = zAMP * Sin(zN*pi + 1)
	
	; Get weighted colors
	fld dword zSinVal
	fld dword z2p
	fmulp st1			; 2.0*zSinVal
	fld1
	faddp st1			; 1 + 2.0*zSinVal
	fld dword zMColWt0
	fmulp st1			; zColorWt(0)*(1 + 2.0*zSinVal)
	fild dword nm255M
	fmulp st1
	fistp dword MiRed	; zRed*255
	
	fld dword zSinVal
	fld dword zp65
	fmulp st1			; 0.65*zSinVal
	fld1
	faddp st1			; 1 + 0.65*zSinVal
	fld dword zMColWt1
	fmulp st1			; zColorWt(1)*(1 + 0.65*zSinVal)
	fild dword nm255M
	fmulp st1
	fistp dword MiGreen	; zGreen*255

	fld dword zSinVal
	fld dword z2p1
	fmulp st1			; 2.1*zSinVal
	fld1
	faddp st1			; 1 + 2.1*zSinVal
	fld dword zMColWt2
	fmulp st1			; zColorWt(2)*(1 + 2.1*zSinVal)
	fild dword nm255M
	fmulp st1
	fistp dword MiBlue	; zBlue*255

	; Make up long color
	mov eax,MiBlue
	and eax,255
	mov ebx,MiGreen
	and ebx,255
	shl ebx,8
	or eax,ebx
	mov ebx,MiRed
	and ebx,255
	shl ebx,16
	or eax,ebx			; eax =  HiB ARGB LoB
	
	stosd				; mov [edi],eax & edi-4

	mov eax,4
	sub esi,eax			; esi-4
	
	dec ecx
	jnz Near MFRIX
	
	dec edx
	jnz Near MFRIY
	
	cld					; Restore direction flag

RET
;=====================================================

SMOOTH:

%define lo32		[ebp-68]
%define hi32		[ebp-72]


	mov eax,PicHeight
	mov ebx,PicWidth
	mul ebx
	shl eax,2			; 4*PicWidth*PicHeight
	mov ecx,eax
	
	mov eax,PicWidth
	shl eax,1			; 2*PicWidth
	inc eax		
	shl eax,2			; 4*(2*PicWidth+1)
	sub ecx,eax			; Num bytes from LongSurf(2,2) to
						; LongSurf(PicWidth,PicHeight-1)
						; NB Top & Bottom rows will not be smoothed
						; and this will prevent several rows from
						; going to one color with repeated smoothing.
						
	mov edi,ptLS		; -> LongSurf(1,1)
	mov eax,PicWidth
	inc eax
	shl eax,2			; 4*(PicWidth+1)
	add edi,eax			; -> LongSurf(2,2)
	
	mov ebx,PicWidth
	shl ebx,2			; 4*PicWidth
	
	mov eax,0FCFCFCFCh	;Make 8 bytes to mask out
	mov lo32,eax		;lower 2 bits so that shr 2
	mov hi32,eax		;divides each byte by 4
	movq mm0,hi32
	
NextPix2:	
	; pick up 4 LongSurf() long integers(4-bytes, 32-bits)
	; on a cross & add together
	movd mm1,[edi-4]
	movd mm2,[edi+4]
	add edi,ebx
	movd mm3,[edi]
	sub edi,ebx
	sub edi,ebx
	movd mm4,[edi]
	add edi,ebx
	
	; AND out 2 lo-bits in each byte so that shift-division 
	; stays in bytes
	pand mm1,mm0
	pand mm2,mm0
	pand mm3,mm0
	pand mm4,mm0
	
	; divide by 4
	psrlq mm1,2
	psrlq mm2,2
	psrlq mm3,2
	psrlq mm4,2
	
	; add them up (unsigned saturation, will go to one color
	; eventually, apart from top & bottom row effect, see above) 
	paddusb mm1,mm2
	paddusb mm3,mm4
	paddusb mm1,mm3
	
	; mov mm1 lo-32-bits into eax
	movd eax,mm1
	
	; put back to cross center
	stosd				; mov [edi],eax  &  edi+4

	sub ecx,4
	jnz Near NextPix2

	emms				;Clear FP/MMX stack

RET
;=====================================================

CYCLE:

	mov eax,PicHeight
	mov ebx,PicWidth
	mul ebx
	shr eax,1			; PicWidth*PicHeight/2  8-byte chunks
	mov ecx,eax
	
	shr ecx,2			; for-32 byte chunks
	
	mov edi,ptLS		; -> LongSurf(1,1)
	
	mov eax,CycleSpeed	; 1,2 or 4
	cmp eax,1
	jne CSp2
	mov eax,01010101h	;Make 8 bytes
	jmp LoadSpeed
CSp2:
	cmp eax,2
	jne CSp4
	mov eax,02020202h	;Make 8 bytes		; double speed
	jmp LoadSpeed
CSp4:
	mov eax,04040404h	;Make 8 bytes		; quad speed

LoadSpeed:
	mov lo32,eax		;to subtract 1,2 or 4 from each byte
	mov hi32,eax
	movq mm4,hi32
	movq mm5,hi32
	movq mm6,hi32
	movq mm7,hi32

	mov ebx,32	;8

NextChunk:
	
	; pick up 32 bytes
	movq mm0,[edi]
	movq mm1,[edi+8]
	movq mm2,[edi+16]
	movq mm3,[edi+24]

	; subtract 1,2 or 4 from each byte with overlap ie 0->255
	psubb mm0,mm4		;sub 8 bytes at a time
	psubb mm1,mm5		;sub 8 bytes at a time
	psubb mm2,mm6		;sub 8 bytes at a time
	psubb mm3,mm7		;sub 8 bytes at a time
	; could do paddb to brighten first
	
	; add/sub 1,2 or 4 from each byte with overlap ie 0->255
	;paddb mm0,mm4		;add 8 bytes at a time
	;psubb mm1,mm5		;sub 8 bytes at a time
	;paddb mm2,mm6		;add 8 bytes at a time
	;psubb mm3,mm7		;sub 8 bytes at a time
	; mixing sub & add produces noise behind a vertical grid pattern

	; Using psubusb or paddusb goes to all black 0,0,0 
	; or all white 255,255,255 min/max unsigned byte
	
	; Using psubsb or paddsb goes to all grey 128,128,128 
	; or all white 255,255,255 min/max signed byte


	; put back 32 bytes
	movq [edi],mm0		;with overlap ie 0->255
	movq [edi+8],mm1	;with overlap ie 0->255
	movq [edi+16],mm2	;with overlap ie 0->255
	movq [edi+24],mm3	;with overlap ie 0->255

	add edi,ebx

	dec ecx
	jnz NextChunk
	emms			;Clear FP/MMX stack

ret

;############################################################
	
	;; Testing ;;;;;;;;
	;mov eax,iBlue
	;cld
	;RET
	;mov ebx,[ebp+12]	; ->zres()
	;mov eax,zr
	;mov [ebx],eax
	;mov eax,zColorWt1
	;mov [ebx+4],eax
	;mov eax,zColorWt2
	;mov [ebx+8],eax
	;mov eax,zv
	;mov [ebx+12],eax
	;cld
	;RET
	;;;;;;;;;;;;;;;;;;;
