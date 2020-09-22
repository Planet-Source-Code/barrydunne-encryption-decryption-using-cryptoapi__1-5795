<div align="center">

## Encryption/Decryption using CryptoAPI


</div>

### Description

This is a module that you can add to your project to encrypt/decrypt using the CryptoAPI. This is the standard API used regardless of who provides the dll which actually does the encryption. Microsoft give you one such dll as standard with windows or NT, but the API ensures that you have the same interface to anyone elses, or even write your own. These different encryption dlls are called Cryptographic Service Providers (CSP's) and the standard Microsoft one is called "Microsoft Base Cryptographic Provider v1.0". To use a different CSP all you have to do is change a constant in this module. This module ensures that there are no carriage returns or line feeds in the encrypted value so that you can easily write it to an ini file for example. This version contains a fix to the original version.
 
### More Info
 
There are two main functions:

'

'Function EncryptData(ByVal Data As String, ByVal Password As String) As String

' Where Data is the String to encrypt and password is used to encrypt it

'

'Function DecryptData(ByVal Data As String, ByVal Password As String) As String

' Where Data is the encrypted String and password is used to decrypt it


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BarryDunne](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/barrydunne.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/barrydunne-encryption-decryption-using-cryptoapi__1-5795/archive/master.zip)





### Source Code

```
Option Explicit
'This is based on the 2 MSDN articles
' "Example C Program: Using CryptAcquireContext"
' "Example C Program: Encrypting a File"
'Example usage:
'
'  Private Const MY_PASSWORD As String = "isdflkaatdfuhwfnasdf"
'
'  Public Sub Main()
'    MsgBox EncryptData("hello world", MY_PASSWORD)
'    MsgBox DecryptData(EncryptData("hello world", MY_PASSWORD), MY_PASSWORD)
'  End Sub
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
  (ByRef phProv As Long, _
   ByVal pszContainer As String, _
   ByVal pszProvider As String, _
   ByVal dwProvType As Long, _
   ByVal dwFlags As Long) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" _
  (ByVal hProv As Long, _
   ByVal dwParam As Long, _
   ByRef pbData As Any, _
   ByRef pdwDataLen As Long, _
   ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
  (ByVal hProv As Long, _
   ByVal Algid As Long, _
   ByVal hKey As Long, _
   ByVal dwFlags As Long, _
   ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
  (ByVal hHash As Long, _
   ByVal pbData As String, _
   ByVal dwDataLen As Long, _
   ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
  (ByVal hProv As Long, _
   ByVal Algid As Long, _
   ByVal hBaseData As Long, _
   ByVal dwFlags As Long, _
   ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
  (ByVal hHash As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" _
  (ByVal hKey As Long, _
   ByVal hHash As Long, _
   ByVal Final As Long, _
   ByVal dwFlags As Long, _
   ByVal pbData As String, _
   ByRef pdwDataLen As Long, _
   ByVal dwBufLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" _
  (ByVal hKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
  (ByVal hProv As Long, _
   ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" _
  (ByVal hKey As Long, _
   ByVal hHash As Long, _
   ByVal Final As Long, _
   ByVal dwFlags As Long, _
   ByVal pbData As String, _
   ByRef pdwDataLen As Long) As Long
Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const KEY_CONTAINER As String = "Metallica"
Private Const PROV_RSA_FULL As Long = 1
Private Const PP_NAME As Long = 4
Private Const PP_CONTAINER As Long = 6
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_MD5 As Long = 3
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Private Const NUMBER_ENCRYPT_PASSWORD As String = "´o¸sçPQ]"
Public Function EncryptData(ByVal Data As String, ByVal Password As String) As String
  Dim sEncrypted As String
  Dim lEncryptionCount As Long
  Dim sTempPassword As String
  'It is possible that the normal encryption will give you a string
  'containing cr or lf characters which make it difficult to write to files
  'Do a loop changing the password and keep encrypting until the result is ok
  'To be able to decrypt we need to also store the number of loops in the result
  'Try first encryption
  lEncryptionCount = 0
  sTempPassword = Password & lEncryptionCount
  sEncrypted = EncryptDecrypt(Data, sTempPassword, True)
  'Loop if this contained a bad character
  Do While (InStr(1, sEncrypted, vbCr) > 0) _
     Or (InStr(1, sEncrypted, vbLf) > 0) _
     Or (InStr(1, sEncrypted, Chr$(0)) > 0) _
     Or (InStr(1, sEncrypted, vbTab) > 0)
    'Try the next password
    lEncryptionCount = lEncryptionCount + 1
    sTempPassword = Password & lEncryptionCount
    sEncrypted = EncryptDecrypt(Data, sTempPassword, True)
    'Don't go on for ever, 1 billion attempts should be plenty
    If lEncryptionCount = 99999999 Then
      Err.Raise vbObjectError + 999, "EncryptData", "This data cannot be successfully encrypted"
      EncryptData = ""
      Exit Function
    End If
  Loop
  'Build encrypted string, starting with number of encryption iterations
  EncryptData = EncryptNumber(lEncryptionCount) & sEncrypted
End Function
Public Function DecryptData(ByVal Data As String, ByVal Password As String) As String
  Dim lEncryptionCount As Long
  Dim sDecrypted As String
  Dim sTempPassword As String
  'When encrypting we may have gone through a number of iterations
  'How many did we go through?
  lEncryptionCount = DecryptNumber(Mid$(Data, 1, 8))
  'start with the last password and work back
  sTempPassword = Password & lEncryptionCount
  sDecrypted = EncryptDecrypt(Mid$(Data, 9), sTempPassword, False)
  DecryptData = sDecrypted
End Function
Public Function GetCSPDetails() As String
  Dim hCryptProv As Long
  Dim lLength As Long
  Dim yContainer() As Byte
  'Get handle to CSP
  If CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
    HandleError "Error during CryptAcquireContext for a new key container." & vbCrLf & _
          "A container with this name probably already exists."
    Exit Function
  End If
  'For developer info, show what the CSP & container name is
  lLength = 1000
  ReDim yContainer(lLength)
  If CryptGetProvParam(hCryptProv, PP_NAME, yContainer(0), lLength, 0) <> 0 Then
    GetCSPDetails = "Cryptographic Service Provider name: " & ByteToStr(yContainer, lLength)
  End If
  lLength = 1000
  ReDim yContainer(lLength)
  If CryptGetProvParam(hCryptProv, PP_CONTAINER, yContainer(0), lLength, 0) <> 0 Then
    GetCSPDetails = GetCSPDetails & vbCrLf & "Key Container name: " & ByteToStr(yContainer, lLength)
  End If
  'Release provider handle.
  If hCryptProv <> 0 Then
    CryptReleaseContext hCryptProv, 0
  End If
End Function
Private Function EncryptDecrypt(ByVal Data As String, ByVal Password As String, ByVal Encrypt As Boolean) As String
  Dim hCryptProv As Long
  Dim lLength As Long
  Dim sTemp As String
  Dim hHash As Long
  Dim hKey As Long
  'Get handle to CSP
  If CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
    If CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0 Then
      HandleError "Error during CryptAcquireContext for a new key container." & vbCrLf & _
            "A container with this name probably already exists."
      Exit Function
    End If
  End If
  '--------------------------------------------------------------------
  'The data will be encrypted with a session key derived from the
  'password.
  'The session key will be recreated when the data is decrypted
  'only if the password used to create the key is available.
  '--------------------------------------------------------------------
  'Create a hash object.
  If CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0 Then
    HandleError "Error during CryptCreateHash!"
  End If
  'Hash the password.
  If CryptHashData(hHash, Password, Len(Password), 0) = 0 Then
    HandleError "Error during CryptHashData."
  End If
  'Derive a session key from the hash object.
  If CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0 Then
    HandleError "Error during CryptDeriveKey!"
  End If
  'Do the work
  sTemp = Data
  lLength = Len(Data)
  If Encrypt Then
    'Encrypt data.
    If CryptEncrypt(hKey, 0, 1, 0, sTemp, lLength, lLength) = 0 Then
      HandleError "Error during CryptEncrypt."
    End If
  Else
    'Encrypt data.
    If CryptDecrypt(hKey, 0, 1, 0, sTemp, lLength) = 0 Then
      HandleError "Error during CryptDecrypt."
    End If
  End If
  'This is what we return.
  EncryptDecrypt = Mid$(sTemp, 1, lLength)
  'Destroy session key.
  If hKey <> 0 Then
    CryptDestroyKey hKey
  End If
  'Destroy hash object.
  If hHash <> 0 Then
    CryptDestroyHash hHash
  End If
  'Release provider handle.
  If hCryptProv <> 0 Then
    CryptReleaseContext hCryptProv, 0
  End If
End Function
Private Sub HandleError(ByVal Error As String)
  'You could write the error to the screen or to a file
  Debug.Print Error
End Sub
Private Function ByteToStr(ByRef ByteArray() As Byte, ByVal lLength As Long) As String
  Dim i As Long
  For i = LBound(ByteArray) To (LBound(ByteArray) + lLength)
    ByteToStr = ByteToStr & Chr$(ByteArray(i))
  Next i
End Function
Private Function EncryptNumber(ByVal lNumber As Long) As String
  Dim i As Long
  Dim sNumber As String
  sNumber = Format$(lNumber, "00000000")
  For i = 1 To 8
    EncryptNumber = EncryptNumber & Chr$(Asc(Mid$(NUMBER_ENCRYPT_PASSWORD, i, 1)) + Val(Mid$(sNumber, i, 1)))
  Next i
End Function
Private Function DecryptNumber(ByVal sNumber As String) As Long
  Dim i As Long
  For i = 1 To 8
    DecryptNumber = (10 * DecryptNumber) + (Asc(Mid$(sNumber, i, 1)) - Asc(Mid$(NUMBER_ENCRYPT_PASSWORD, i, 1)))
  Next i
End Function
```

