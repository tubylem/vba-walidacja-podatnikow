''
' vba-walidacja-podatnikow
' (c) Robert Soszynski - https://github.com/tubylem/vba-walidacja-podatnikow
'
' Checking taxpayers based on the NIP number using the WL Register API
'
' @class walidacja-podatnikow
' @author robert.soszynski@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Function GetRequest(Url As String) As String

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    
    GetRequest = objHTTP.responseText
    
End Function

Private Function CleanNumber(number As String) As String
    Dim ch, bytes() As Byte: bytes = number
    For Each ch In bytes
        If Chr(ch) Like "[0-9]" Then CleanNumber = CleanNumber & Chr(ch)
    Next ch
End Function

Private Function GetAttribute(NIP As String, name As String) As String

    NIP = CleanNumber(NIP)

    Dim Message As String
    Dim Url As String
    
    Url = "https://wl-api.mf.gov.pl/api/search/nips/" & NIP & "?date=" & Format(Date, "yyyy-mm-dd")
    
    JsonString = GetRequest(Url)
    Set jsonObject = JsonConverter.ParseJson(JsonString)
    
    If jsonObject("code") <> "" Then
        Message = "BŁĄD: " & jsonObject("message")
    ElseIf jsonObject("result")("subjects").Count = 0 Then
        Message = "BŁĄD: " & "Podatnik nie istnieje"
    Else
        For Each Item In jsonObject("result")("subjects")
            If TypeName(Item(name)) = "String" Then
                Message = Item(name)
            ElseIf TypeName(Item(name)) = "Null" Then
                Message = "Brak danych"
            ElseIf TypeName(Item(name)) = "Collection" Then
                For Each Value In Item(name)
                    Message = Message & "," & Chr(13) & Value
                Next Value
                Message = Mid(Message, 3)
            Else
                Message = JsonConverter.ConvertToJson(Item(name))
            End If
        Next
    End If
    
    GetAttribute = Message
    
End Function

Function PodmiotStatus(NIP As String) As String

    PodmiotStatus = GetAttribute(NIP, "statusVat") & " [" & Date & "]"
    
End Function

Function PodmiotNazwa(NIP As String) As String

    PodmiotNazwa = GetAttribute(NIP, "name")
    
End Function

Function PodmiotRegon(NIP As String) As String

    PodmiotRegon = GetAttribute(NIP, "regon")
    
End Function

Function PodmiotPesel(NIP As String) As String

    PodmiotPesel = GetAttribute(NIP, "pesel")
    
End Function

Function PodmiotKrs(NIP As String) As String

    PodmiotKrs = GetAttribute(NIP, "krs")
    
End Function

Function PodmiotDataRejestracji(NIP As String) As String

    PodmiotDataRejestracji = GetAttribute(NIP, "registrationLegalDate")
    
End Function

Function PodmiotKontaBankowe(NIP As String) As String

    PodmiotKontaBankowe = GetAttribute(NIP, "accountNumbers")

End Function