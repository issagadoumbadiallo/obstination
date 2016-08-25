
'##########################################################################################################
'##########################################################################################################
'#
'#  Fonction qui sera appelée pour effectuer la tache de mise à jour du Pré-Trade en fonction du Post-Trade
'#  Date : 22/10/2014
'#  proc stoc = "compliance_update_Rev4_CS_TEST_V2"
'##########################################################################################################
'##########################################################################################################
Public Function fCTRL_MAJ_OLD_Ctre_Pretrade() As Boolean

'Initialisation de la proc stock "compliance_update_Rev4_CS_TEST_V2"
Set ps = psTEST_V2

'rs old name = rst_CTRL_AUTO_MAJ_OLD_CONTRAINTES

Set rs = CurrentDb.OpenRecordset("PRETRADE_R_CTRL_AUTO_MAJ_OLD_CONTRAINTES")
If rs.EOF = True And rs.BOF = False Then rs.MoveFirst

        While rs.EOF = False
'Initialisation des paramètres de la proc stock ps(TRUE)
ps(True).Parameters("@TEST_ID") = rs.Fields("TEST_ID_PRE_TRADE")

'###########################################
'###########################################
'#  Gestion des contraintes d'interdiction
'###########################################
'###########################################
If ps.Parameters("@TEST_ID") = "EXC" Then
'----- EXPRESSION -----
ps.Parameters("@EXPRESSION").Value = "( " & rs.Fields("EXPRESSION_POST_TRADE") & " ) and ORDTRANS = BUYL"
        '----- CASE Portfolio et Pré-Trade -----
ps.Parameters("@BATCH_MODE").Value = "N"
ps.Parameters("@PRE_TRADE_MODE").Value = "Y"
        '----- Severity (Warning "S" et non Alert "H") -----
sSeverity = "S"
ps.Parameters("@NUMERATOR_CD").Value = ""
ps.Parameters("@DENOMINATOR_CD").Value = ""

        '###########################################
'###########################################
'#  Gestion des contraintes calculatoire
'###########################################
'###########################################
ElseIf ps.Parameters("@TEST_ID") = "CON" Then

'----- EXPRESSION -----
ps.Parameters("@EXPRESSION").Value = "( " & myEXPRESSION_POST_TRADE & " ) and ATTRADETRANS = ALL"

        '----- CASE Portfolio et Pré-Trade -----
ps.Parameters("@BATCH_MODE").Value = "Y"
ps.Parameters("@PRE_TRADE_MODE").Value = "Y"

        '----- Changement de bornes pour le Calculatoire -----
ps.Parameters("@UPPER_LIMIT") = Null
        ps.Parameters("@LOWER_LIMIT") = Null
        ps.Parameters("@UPPER_WARN") = rs.Fields("UPPER_LIMIT")
ps.Parameters("@LOWER_WARN") = rs.Fields("LOWER_LIMIT")

'----- Numérateur et Dénominateur -----
If rs.Fields("NUMERATOR_CD") = "POSNUMLB7" Or rs.Fields("NUMERATOR_CD") = "POSNUMLB9" Then
        ps.Parameters("@NUMERATOR_CD").Value = "MKT_VAL"
ElseIf rs.Fields("NUMERATOR_CD") = "POSNUMLB8" Then
        ps.Parameters("@NUMERATOR_CD").Value = "QTY"
End If
ps.Parameters("@DENOMINATOR_CD").Value = "TOT_ASSET"

        '###########################################
'###########################################
'#  Gestion des contraintes comptage
'###########################################
'###########################################
ElseIf ps.Parameters("@TEST_ID") = "BKT" Then
'----- EXPRESSION -----
ps.Parameters("@EXPRESSION").Value = "( " & myEXPRESSION_POST_TRADE & " ) and ATTRADETRANS = ALL"
        '----- CASE Portfolio et Pré-Trade -----
ps.Parameters("@BATCH_MODE").Value = "Y"
ps.Parameters("@PRE_TRADE_MODE").Value = "Y"
        '----- Changement de bornes pour le Comptage -----
iUPPER_LIMIT = Null
iLOWER_LIMIT = Null
iUPPER_WARN = myUPPER_LIMIT_POST_TRADE
iLOWER_WARN = myLOWER_LIMIT_POST_TRADE
'----- Paramètre du compteur -----
sFOR_EACH_PARAM_1 = myFOR_EACH_PARAM_1
sFOR_EACH_PARAM_2 = myFOR_EACH_PARAM_2
'###########################################
'###########################################
'#  Gestion des contraintes de restriction
'###########################################
'###########################################
ElseIf ps.Parameters("@TEST_ID") = "VAL" Then
'----- EXPRESSION -----
ps.Parameters("@EXPRESSION").Value = "( " & myEXPRESSION_POST_TRADE & " ) and ATTRADETRANS = ALL"
        '----- CASE Portfolio et Pré-Trade -----
ps.Parameters("@BATCH_MODE").Value = "Y"
ps.Parameters("@PRE_TRADE_MODE").Value = "Y"
        '----- Changement de bornes pour le Calculatoire -----
iUPPER_LIMIT = Null
iLOWER_LIMIT = Null
iUPPER_WARN = myUPPER_LIMIT_POST_TRADE
iLOWER_WARN = myLOWER_LIMIT_POST_TRADE
'----- Numérateur et Dénominateur -----
If rs.Fields("NUMERATOR_CD") = "POSNUMLB7" Or rs.Fields("NUMERATOR_CD") = "POSNUMLB9" Then
        ps.Parameters("@NUMERATOR_CD").Value = "MKT_VAL"
ElseIf rs.Fields("NUMERATOR_CD") = "POSNUMLB8" Then
        ps.Parameters("@NUMERATOR_CD").Value = "QTY"
End If
ps.Parameters("@DENOMINATOR_CD").Value = "TOT_ASSET"
End If

'###########################################
'###########################################
'#  Modifications communes
'###########################################
'###########################################
'----- VIOL_NOTE -----
myTempText = "Ce dépassement de type 'Warning' n'est pas bloquant, l'ordre a été transmis à la négociation (sauf si vous l'aviez simplement simulé). Pour annuler votre ordre, appelez le négociateur."
myTempText = myTempText & vbCrLf
myTempText = myTempText & "Vous pouvez aussi contacter les Risk Managers RCO au 49777."
sViolNote = myTempText

'----- TEST_NAME : Préfixe (Pré-Trade) -----
stestname = "(Pré-Trade) " & myTEST_NAME_POST_TRADE
'----- Information relative à la modification (utilisateur = "CTRLRISK", date = Date de la MaJ) -----
sRevBy = "CTRLRISK"
sDescription = myDESCRIPTION
sDate = Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & Hour(Time) & ":" & Minute(Time)

'----- INACTIVTD_DATE / INACTIVTD_BY / INACTIVTD_FLAG -----

If myINACTIVTD_FLAG = "Y" Then
        sInactive = myINACTIVTD_FLAG
sInactiveBy = myINACTIVTD_BY
sInactiveDate = myINACTIVTD_DATE
Cmd.Parameters("@INACTVTD_FLAG").Value = sInactive
Cmd.Parameters("@INACTIVTD_BY").Value = sInactiveBy
Cmd.Parameters("@INACTVTD_DATE").Value = sDate
ElseIf myINACTIVTD_FLAG = "N" Then
        Cmd.Parameters("@INACTVTD_FLAG").Value = sInactive
Cmd.Parameters("@INACTIVTD_BY").Value = ""
Cmd.Parameters("@INACTVTD_DATE").Value = Null
End If

Cmd.Parameters("@TEST_ID").Value = myTEST_ID_PRE_TRADE         '@TestID
Cmd.Parameters("@BATCH_MODE").Value = ps.Parameters("@BATCH_MODE").Value               '@Batch mode (case Post Trade J-1)
Cmd.Parameters("@PRE_TRADE_MODE").Value = ps.Parameters("@PRE_TRADE_MODE").Value            '@Pre trade mode (case Pre trade)
Cmd.Parameters("@SEVERITY").Value = sSeverity                  '@SEVERITY (Warning [S] ou Alerte [H])
Cmd.Parameters("@EXPRESSION").Value = ps.Parameters("@EXPRESSION").Value              '@EXPRESSION (Where Clause)
Cmd.Parameters("@VIOL_NOTE").Value = sViolNote                 '@VIOL_NOTE (Message d'alerte)
Cmd.Parameters("@TEST_NAME").Value = Fit_String(stestname)     '@TEST_NAM (Libellé Contrainte)
Cmd.Parameters("@DESCRIPTION").Value = sDescription            '@DESCRIPTION (Description de la contrainte)
Cmd.Parameters("@REVSD_BY").Value = sRevBy                     '@REV_BY (champ USER derniere sauvegarde)
Cmd.Parameters("@LAST_REVSD_DATE").Value = sDate               '@LAST_REVSD_DATE (champ DATE derniere sauvegarde  au format "mm/dd/yyy")
Cmd.Parameters("@NUMERATOR_CD").Value = ps.Parameters("@NUMERATOR_CD").Value           '@NUMERATOR_CD (champ Numerateur)
Cmd.Parameters("@DENOMINATOR_CD").Value = ps.Parameters("@DENOMINATOR_CD").Value       '@DENOMINATOR_CD (champ Numerateur)
Cmd.Parameters("@UPPER_LIMIT").Value = iUPPER_LIMIT            '@UPPER_LIMIT (champ Borne Alerte Max)
Cmd.Parameters("@UPPER_WARN").Value = iUPPER_WARN              '@UPPER_WARN (champ Borne Warning Max)
Cmd.Parameters("@LOWER_LIMIT").Value = iLOWER_LIMIT            '@LOWER_LIMIT (champ Borne Alerte Min)
Cmd.Parameters("@LOWER_WARN").Value = iLOWER_WARN              '@LOWER_WARN (champ Borne Warning Min)
'----- En attente que la procédure stockée soit mise à jour -----
'        Cmd.Parameters(20).Value = sFOR_EACH_PARAM_1                  '@FOR_EACH_PARAM_1 (champ Paramètre du compteur)
'        Cmd.Parameters(21).Value = sFOR_EACH_PARAM_2                  '@FOR_EACH_PARAM_2 (champ Paramètre du compteur)
'        Cmd.Parameters(22).Value = sFOR_EACH_VAL_1                    '@FOR_EACH_VAL_1 (champ Paramètre du compteur)
'        Cmd.Parameters(23).Value = sFOR_EACH_VAL_2                    '@FOR_EACH_VAL_2 (champ Paramètre du compteur)

Cmd.Execute

        rs.MoveNext
        Wend



'#######################################################################
'# Gestion des contrainte de type "Restric the Value" (TEST_TYPE = VAL")
'#######################################################################
'Set rs_RESTRICT_VAL = CurrentDb.OpenRecordset("PRETRADE_R_CTRL_AUTO_MAJ_OLD_CONTRAINTES_RESTRICT_VALUE")
'If rs_RESTRICT_VAL.EOF = True And rs_RESTRICT_VAL.BOF = False Then rs_RESTRICT_VAL.MoveFirst

'While rst_CTRL_AUTO_MAJ_NEW_CONTRAINTES_RESTRICT_VAL.EOF = False



'Wend

'Initialisation de la commande
Set Cmd = Nothing

End Function
