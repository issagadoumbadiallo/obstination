Attribute VB_Name = "AAAA"
Option Compare Database
Option Explicit

Sub ssf()

    Dim sPseudo As String
    
    Dim rst As Recordset2
    
    Set rst = CurrentDb.OpenRecordset("SELECT PSEUDO, LINK FROM T_JOUEUR AS T WHERE T.COMPUTE=True")
    
    Dim sUrlToCall As String
    
    While Not rst.EOF
        
        sUrlToCall = rst("LINK").Value
                        

        rst.MoveNext
    Wend
    
End Sub


Function ConformizeData()
    Dim rs As Recordset2
    Set rs = CurrentDb.OpenRecordset("T_COTE_RESULT")
    
    While Not rs.EOF
        rs.Edit
            
            rs("NM_RENCONTRE").Value = Replace(rs("NM_RENCONTRE").Value, "Etats-Unis", "Etats_Unis")
            rs("NM_RENCONTRE").Value = Replace(rs("NM_RENCONTRE").Value, "Pays-Bas", "Pays_Bas")
            rs("NM_RENCONTRE").Value = Replace(rs("NM_RENCONTRE").Value, "Hammam-Lif", "Hammam_Lif")
            rs("NM_RENCONTRE").Value = Replace(rs("NM_RENCONTRE").Value, "Burkina-Faso", "Burkina_Faso")
            
        rs.update
        
        rs.MoveNext
    Wend
    
'Check if Error
'SELECT T_COTE_RESULT.ID, T_COTE_RESULT.NM_RENCONTRE, fMatch([NM_RENCONTRE],1) AS Equip1, fMatch([NM_RENCONTRE],2) AS Equip2
'FROM T_COTE_RESULT
'WHERE (((fMatch([NM_RENCONTRE],1)) Like "ERROR")) OR (((fMatch([NM_RENCONTRE],2)) Like "ERROR"));

End Function


Function fMatch(ByVal Rencontre As String, ordre As Integer) As String
    Dim Tsplit
    Dim bGetIn As Boolean
    
    Tsplit = Split(Rencontre, " - ") 'Traitement 2 tirets
    
    If UBound(Tsplit) = 1 And Not bGetIn Then
        fMatch = OptimiseForLevDist(CStr(Tsplit(ordre - 1)))
        bGetIn = True
    End If
    
    
    Tsplit = Split(Rencontre, "-") 'Traitement 1 tiret
    If UBound(Tsplit) = 1 And Not bGetIn Then
        fMatch = OptimiseForLevDist(CStr(Tsplit(ordre - 1)))
        bGetIn = True
    End If
    
    If Not bGetIn Then
        fMatch = "ERROR"
    End If

End Function

Sub DoneRoutine()

'R_INFO_BY_TEAM
CurrentDb.Execute "SELECT T_COTE_RESULT.ID, T_COTE_RESULT.NM_RENCONTRE, T_COTE_RESULT.TOURNOI, T_COTE_RESULT.COTE1, T_COTE_RESULT.COTEN, T_COTE_RESULT.COTE2, T_COTE_RESULT.SCORE, T_COTE_RESULT.SCORE1, T_COTE_RESULT.SCOREN, T_COTE_RESULT.SCORE2, T_COTE_RESULT.RESULT, T_COTE_RESULT.PARIPRONO, T_COTE_RESULT.DT_MATCH, T_COTE_RESULT.DT_INS, T_COTE_RESULT.FICHEENCOURS, fMatch([NM_RENCONTRE],1) AS Equip, 1 AS TEAM" & vbCr & _
"FROM T_COTE_RESULT" & vbCr & _
"UNION ALL" & vbCr & _
"SELECT T_COTE_RESULT.ID, T_COTE_RESULT.NM_RENCONTRE, T_COTE_RESULT.TOURNOI, T_COTE_RESULT.COTE1, T_COTE_RESULT.COTEN, T_COTE_RESULT.COTE2, T_COTE_RESULT.SCORE, T_COTE_RESULT.SCORE1, T_COTE_RESULT.SCOREN, T_COTE_RESULT.SCORE2, T_COTE_RESULT.RESULT, T_COTE_RESULT.PARIPRONO, T_COTE_RESULT.DT_MATCH, T_COTE_RESULT.DT_INS, T_COTE_RESULT.FICHEENCOURS, fMatch([NM_RENCONTRE],2) AS Equip, 2 AS TEAM" & vbCr & _
"FROM T_COTE_RESULT"
'

'Select only needed column form R_INFO_BY_TEAM
CurrentDb.Execute "SELECT R_INFO_BY_TEAM.* INTO T_RENCONTRE" & vbCr & _
"FROM R_INFO_BY_TEAM;"

'Traiter le nom des tournois
'Liste le nom distinct des tournois puis uniformiser les noms
'Use T_Rencontre and rows you need to be computed
'


'First Give a notation (1 3 5) ponderated by the league - give a score to the league while normalizing leagues name
'Cumulate notation

'Find a way to say if
'1- Bonne dynamique
'2- Ready to break - tied
'3- Going to play safe
'4- Have to win - no choice
'5- not trying anymore - relegation

End Sub


