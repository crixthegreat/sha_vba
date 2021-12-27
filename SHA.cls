VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' SHA-1 & 256 algorithm for VBA by Crix
' based on Federal Information Processing Standard (FIPS) 180-4
' import to class module to use
' 2021-10-28 ~ 2021-12-27

Private BLOCK_LENGTH As Integer  '64 bytes = 512 bits

Private HASH() As Long, K() As Long, W(80) As Long, sum_K As Byte
Private len_m As Long, block_sum As Long
Private i As Long, j As Long, n As Long
Private message_buff() As Byte
Private len_m_hex As String
Private M() As Byte, sub_block_str As String
Private HASH_SEGS As Byte


Public Function sha_string(ByVal text As String, Optional ByVal SHA As String = "SHA-1") As String
    Dim last_data_block As Long
    
    'read the message into bytes
    message_buff = StrConv(text, vbFromUnicode)
    len_m = Len(text)

    Select Case SHA
        Case "SHA-1"
            BLOCK_LENGTH = 64
            HASH_SEGS = 5
            sum_K = 80
        Case "SHA-256"
            BLOCK_LENGTH = 64
            HASH_SEGS = 8
            sum_K = 64
        Case Else
            MsgBox "unknown SHA type when reading message: " & SHA
            End
    End Select
    
    block_sum = CLng(len_m / BLOCK_LENGTH)
    ReDim HASH(block_sum + 1, HASH_SEGS - 1) As Long
    ReDim K(sum_K - 1) As Long

    '____________STEP 1
    'Padding the Message and parsing into N m-bit blocks
    
    PaddingMessage (SHA)

    '____________STEP 2
    'Initial the HASH() and constant K()

    InitialHashAndK (SHA)
    
    '____________STEP 3
    'Calculate the HASH
    CalculateHash (SHA)
    
    
    sha_string = ""
    For i = 0 To HASH_SEGS - 1
        Debug.Print (Hex(HASH(block_sum, i)))
        sha_string = sha_string & String(8 - Len(Hex(HASH(block_sum, i))), "0") & Hex(HASH(block_sum, i))
    Next
End Function

Private Sub PaddingMessage(Optional ByVal SHA As String = "SHA-1")
    If len_m Mod 64 <= 56 Then
        block_sum = block_sum + 1
    Else
        block_sum = block_sum + 2
    End If
    
    'parse into blocks

    ReDim M(block_sum, BLOCK_LENGTH) As Byte
    
    'read the message to M
    n = 0
    For i = 0 To block_sum - 1
        For j = 0 To BLOCK_LENGTH - 1
            If n = len_m Then Exit For
            M(i, j) = message_buff(n)
            n = n + 1
        Next
    Next
    
    'set the last bit to 1
    M(Int(len_m / 64), len_m Mod 64) = &H80 '&H80 = 1000 0000

    'set the length of the message
    i = 1
    len_m_hex = Hex(len_m * 8) 'length by bit
    'Debug.Print len_m_hex
    If Len(len_m_hex) Mod 2 <> 0 Then len_m_hex = Space(1) & len_m_hex
    
    Do
        M(block_sum - 1, BLOCK_LENGTH - i) = CInt("&H" & Trim(Mid(len_m_hex, Len(len_m_hex) - i * 2 + 1, 2)))
        i = i + 1
    Loop Until i * 2 > Len(len_m_hex)

    
    For i = 0 To block_sum - 1
        For j = 0 To BLOCK_LENGTH - 1
            Cells(3, 2).Offset(Int(j / 8), j Mod 8 + i * 8).Value = M(i, j)
        Next
    Next
End Sub

Private Sub InitialHashAndK(Optional ByVal SHA As String = "SHA-1")

    Select Case SHA
        Case "SHA-1"
            HASH(0, 0) = &H67452301
            HASH(0, 1) = &HEFCDAB89
            HASH(0, 2) = &H98BADCFE
            HASH(0, 3) = &H10325476
            HASH(0, 4) = &HC3D2E1F0
            
            For i = 0 To sum_K - 1
                Select Case i
                    Case 0 To 19
                        K(i) = &H5A827999
                    Case 20 To 39
                        K(i) = &H6ED9EBA1
                    Case 40 To 59
                        K(i) = &H8F1BBCDC
                    Case 60 To 79
                        K(i) = &HCA62C1D6
                    Case Else
                        
                End Select
            Next
        Case "SHA-256"
            HASH(0, 0) = &H6A09E667
            HASH(0, 1) = &HBB67AE85
            HASH(0, 2) = &H3C6EF372
            HASH(0, 3) = &HA54FF53A
            HASH(0, 4) = &H510E527F
            HASH(0, 5) = &H9B05688C
            HASH(0, 6) = &H1F83D9AB
            HASH(0, 7) = &H5BE0CD19
            
            K(0) = &H428A2F98
            K(1) = &H71374491
            K(2) = &HB5C0FBCF
            K(3) = &HE9B5DBA5
            K(4) = &H3956C25B
            K(5) = &H59F111F1
            K(6) = &H923F82A4
            K(7) = &HAB1C5ED5
            K(8) = &HD807AA98
            K(9) = &H12835B01
            K(10) = &H243185BE
            K(11) = &H550C7DC3
            K(12) = &H72BE5D74
            K(13) = &H80DEB1FE
            K(14) = &H9BDC06A7
            K(15) = &HC19BF174
            K(16) = &HE49B69C1
            K(17) = &HEFBE4786
            K(18) = &HFC19DC6
            K(19) = &H240CA1CC
            K(20) = &H2DE92C6F
            K(21) = &H4A7484AA
            K(22) = &H5CB0A9DC
            K(23) = &H76F988DA
            K(24) = &H983E5152
            K(25) = &HA831C66D
            K(26) = &HB00327C8
            K(27) = &HBF597FC7
            K(28) = &HC6E00BF3
            K(29) = &HD5A79147
            K(30) = &H6CA6351
            K(31) = &H14292967
            K(32) = &H27B70A85
            K(33) = &H2E1B2138
            K(34) = &H4D2C6DFC
            K(35) = &H53380D13
            K(36) = &H650A7354
            K(37) = &H766A0ABB
            K(38) = &H81C2C92E
            K(39) = &H92722C85
            K(40) = &HA2BFE8A1
            K(41) = &HA81A664B
            K(42) = &HC24B8B70
            K(43) = &HC76C51A3
            K(44) = &HD192E819
            K(45) = &HD6990624
            K(46) = &HF40E3585
            K(47) = &H106AA070
            K(48) = &H19A4C116
            K(49) = &H1E376C08
            K(50) = &H2748774C
            K(51) = &H34B0BCB5
            K(52) = &H391C0CB3
            K(53) = &H4ED8AA4A
            K(54) = &H5B9CCA4F
            K(55) = &H682E6FF3
            K(56) = &H748F82EE
            K(57) = &H78A5636F
            K(58) = &H84C87814
            K(59) = &H8CC70208
            K(60) = &H90BEFFFA
            K(61) = &HA4506CEB
            K(62) = &HBEF9A3F7
            K(63) = &HC67178F2

        Case Else
            MsgBox "unknown HASH type when initial contanst: " & SHA
            End
    End Select

    
    
End Sub

Private Function Lshift(ByVal var As Long, ByVal shift As Byte) As Long
    
    'Lshift function for 32bit number
    'Function from https://www.excely.com/excel-vba/bit-shifting-function.shtml
   
    Dim i As Byte, M As Long
    Lshift = var
    
    If shift > 0 Then
        For i = 1 To shift
            M = Lshift And &H40000000 'save 30th bit
            Lshift = (Lshift And &H3FFFFFFF) * 2  ' clear the 30th and the 31st bits, and shift
            If M <> 0 Then
                Lshift = Lshift Or &H80000000 'set the 32th bit
            End If
        Next i
    End If
End Function

Private Function Rshift(ByVal var As Long, ByVal shift As Byte) As Long
    Dim M As Long
    
    Rshift = var
    If shift = 0 Then Exit Function

    M = Rshift And &H80000000 'save the 32th bit
    
    Rshift = Rshift And &H7FFFFFFF 'clear the 32th bit
    Rshift = Int(Rshift / (2 ^ shift)) 'shift right
    
    If M <> 0 Then  'shift the most left bit (sign bit)
        Rshift = Rshift + 2 ^ (31 - shift)
    End If
End Function

Private Function ROTL(ByVal var As Long, shift As Byte) As Long
    ROTL = Lshift(var, shift) Or Rshift(var, 32 - shift)
End Function

Private Function ROTR(ByVal var As Long, shift As Byte) As Long
    ROTR = Rshift(var, shift) Or Lshift(var, 32 - shift)
End Function

Private Function f(ByVal t As Byte, ByVal b As Long, ByVal c As Long, ByVal d As Long)

    ' function f(t, b, c, d) for sha-1
    Select Case t
        Case 0 To 19
            f = Ch(b, c, d)
        Case 20 To 39
            f = Parity(b, c, d)
        Case 40 To 59
            f = Maj(b, c, d)
        Case 60 To 79
            f = Parity(b, c, d)
        Case Else
            MsgBox "wrong value of t= " & Str(t)
    End Select
End Function

Private Function Sigma_U_0(ByVal x As Long)
    ' used in sha-256, sha-224
    Sigma_U_0 = ROTR(x, 2) Xor ROTR(x, 13) Xor ROTR(x, 22)
End Function

Private Function Sigma_U_1(ByVal x As Long)
    ' used in sha-256, sha-224
    Sigma_U_1 = ROTR(x, 6) Xor ROTR(x, 11) Xor ROTR(x, 25)
End Function

Private Function Sigma_L_0(ByVal x As Long)
    ' used in sha-256, sha-224
    Sigma_L_0 = ROTR(x, 7) Xor ROTR(x, 18) Xor Rshift(x, 3)
End Function

Private Function Sigma_L_1(ByVal x As Long)
    ' used in sha-256, sha-224
    Sigma_L_1 = ROTR(x, 17) Xor ROTR(x, 19) Xor Rshift(x, 10)
End Function

Private Function Ch(x As Long, y As Long, z As Long)
    Ch = (x And y) Xor ((Not x) And z)
End Function

Private Function Parity(x As Long, y As Long, z As Long)
    Parity = x Xor y Xor z
End Function

Private Function Maj(x As Long, y As Long, z As Long)
    Maj = (x And y) Xor (x And z) Xor (y And z)
End Function

Private Function Amod(ByVal x As Long, ByVal y As Long) As Long
    Dim i As Byte, carry_flag As Boolean, bit_value As Boolean, x_bit As Boolean, y_bit As Boolean
    
    Amod = 0
    carry_flag = False
    bit_value = False
    For i = 1 To 31
        x_bit = getbit(x, i)
        y_bit = getbit(y, i)
        bit_value = x_bit Xor y_bit Xor carry_flag
        carry_flag = (x_bit And y_bit) Or ((x_bit Or y_bit) And carry_flag)
        If bit_value Then Amod = Amod + 2 ^ (i - 1)
    Next
    
    x_bit = getbit(x, 32)
    y_bit = getbit(y, 32)
    bit_value = x_bit Xor y_bit Xor carry_flag

    If bit_value Then Amod = Amod - 2 ^ 31
End Function

Private Function getbit(ByVal val As Long, ByVal bit As Byte) As Boolean
    If bit < 1 Then Exit Function
    If bit = 32 Then
        If val < 0 Then getbit = True Else getbit = False
        Exit Function
    End If
    
    If (val And 2 ^ (bit - 1)) = 0 Then
        getbit = False
    Else
        getbit = True
    End If
End Function

Private Sub CalculateHash(Optional ByVal SHA As String = "SHA-1")
    Dim a_ As Long, b_ As Long, c_ As Long, d_ As Long, e_ As Long, t As Long, t1 As Long
    Dim f_ As Long, g_ As Long, h_ As Long
    Dim r As Integer, byte_str(4) As String
    Dim sub_block(16) As Long

    For i = 0 To block_sum - 1
    
        For j = 0 To HASH_SEGS - 1
            With Cells(112, 3)
                .Offset(i, j).Value = Hex(HASH(i, j))
            End With
        Next
                
        
        ' divide M(i) into 16 4-bytes-sub_blocks
        ' shared in SHA-1, SHA-256
        For j = 0 To 15
            sub_block_str = ""
            For r = 0 To 3
                byte_str(r) = Hex(M(i, j * 4 + r))
                If Len(byte_str(r)) = 1 Then byte_str(r) = "0" & byte_str(r)
                sub_block_str = sub_block_str & byte_str(r)
            Next
            sub_block(j) = CLng("&H" & sub_block_str)
            
            Cells(12, 2).Offset(Int(j / 8), j Mod 8).Value = Hex(sub_block(j))
        Next
        
        ' calculate W(t)
        For j = 0 To sum_K - 1
            Select Case SHA
                Case "SHA-1"
                    ' calculate W(t), expand the message blocks to 320 bytes
                    ' for SHA-1
                    If j <= 15 Then
                        W(j) = sub_block(j)
                    Else
                        W(j) = ROTL(W(j - 3) Xor W(j - 8) Xor W(j - 14) Xor W(j - 16), 1)
                    End If
                Case "SHA-256"
                    ' calculate W(t) for SHA-256
                    If j <= 15 Then
                        W(j) = sub_block(j)
                    Else
                        W(j) = Amod(Amod(Amod(Sigma_L_1(W(j - 2)), W(j - 7)), Sigma_L_0(W(j - 15))), W(j - 16))
                    End If
                Case Else
                    MsgBox "unknown HASH type when calculate W(t): " & SHA
                    End
            End Select
            
            Cells(15, 2).Offset(Int(j / 8), j Mod 8).Value = Hex(W(j))
        Next

        ' caluculate working variables
        ' calculate a,b,c,d,e (for sha-1)
        a_ = HASH(i, 0)
        b_ = HASH(i, 1)
        c_ = HASH(i, 2)
        d_ = HASH(i, 3)
        e_ = HASH(i, 4)
        
        Select Case SHA
            Case "SHA-1"
                ' need do nothing here
            Case "SHA-256"
                ' and f,g,h (for sha-256)
                f_ = HASH(i, 5)
                g_ = HASH(i, 6)
                h_ = HASH(i, 7)
            Case Else
                MsgBox "unknown HASH type when calculate a,b,c,d,e,f,g,h: " & SHA
                End
        End Select

        For j = 0 To sum_K - 1
            With Cells(28, 3)
                .Offset(j, 0) = Hex(a_)
                .Offset(j, 1) = Hex(b_)
                .Offset(j, 2) = Hex(c_)
                .Offset(j, 3) = Hex(d_)
                .Offset(j, 4) = Hex(e_)
            End With
            
            Select Case SHA
                Case "SHA-1"
                        t = Amod(Amod(Amod(Amod(ROTL(a_, 5), f(j, b_, c_, d_)), e_), K(j)), W(j))
                        e_ = d_
                        d_ = c_
                        c_ = ROTL(b_, 30)
                        b_ = a_
                        a_ = t
                Case "SHA-256"
                        With Cells(28, 3)
                            .Offset(j, 5) = Hex(f_)
                            .Offset(j, 6) = Hex(g_)
                            .Offset(j, 7) = Hex(h_)
                        End With
                        t = Amod(Amod(Amod(Amod(h_, Sigma_U_1(e_)), Ch(e_, f_, g_)), K(j)), W(j))
                        t1 = Amod(Sigma_U_0(a_), Maj(a_, b_, c_))
                        h_ = g_
                        g_ = f_
                        f_ = e_
                        e_ = Amod(d_, t)
                        d_ = c_
                        c_ = b_
                        b_ = a_
                        a_ = Amod(t, t1)
                Case Else
                    MsgBox "unknown HASH type when calculate a,b,c,d,e,f,g,h: " & SHA
                    End
            End Select
        Next

        
        With Cells(28, 3)
            .Offset(80, 0) = Hex(a_)
            .Offset(80, 1) = Hex(b_)
            .Offset(80, 2) = Hex(c_)
            .Offset(80, 3) = Hex(d_)
            .Offset(80, 4) = Hex(e_)
        End With
        
        ' last step, generate hash code
        HASH(i + 1, 0) = Amod(a_, HASH(i, 0))
        HASH(i + 1, 1) = Amod(b_, HASH(i, 1))
        HASH(i + 1, 2) = Amod(c_, HASH(i, 2))
        HASH(i + 1, 3) = Amod(c_, HASH(i, 3))
        HASH(i + 1, 4) = Amod(e_, HASH(i, 4))
        
        Select Case SHA
            Case "SHA-1"

            Case "SHA-256"
                HASH(i + 1, 5) = Amod(f_, HASH(i, 5))
                HASH(i + 1, 6) = Amod(g_, HASH(i, 6))
                HASH(i + 1, 7) = Amod(h_, HASH(i, 7))
            Case Else
                MsgBox "unknown HASH type when generate hash code: " & SHA
                End
        End Select
    Next
End Sub
