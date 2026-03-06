Attribute VB_Name = "MailsLOD1"

Option Explicit

' ============================================================
'  CONFIGURATION — adapter avant utilisation
' ============================================================
Private Const EXPEDITEUR_NOM   As String = "Gianni"
Private Const EXPEDITEUR_EMAIL As String = "gianni.xxx@amundi.com"
Private Const DIRECTION        As String = "Opérations & Conformité Externalisations — Amundi Immobilier"

' Colonnes DATABASE (1-based)
Private Const COL_NOM       As Integer = 1   ' A - Nom prestataire
Private Const COL_TYPE      As Integer = 5   ' E - PCI/PS
Private Const COL_INTRAHG   As Integer = 6   ' F - Intra/HG
Private Const COL_RESP_NOM  As Integer = 11  ' K - Nom responsable interne
Private Const COL_RESP_MAIL As Integer = 12  ' L - Email responsable interne
Private Const COL_PCA_NOM   As Integer = 13  ' M - Nom contact PCA prestataire
Private Const COL_PCA_MAIL  As Integer = 14  ' N - Email contact PCA prestataire

' Colonnes CONTROLES LOD1 (1-based)
Private Const COL_CTRL_NOM  As Integer = 1   ' A - Nom (repris de DATABASE)
Private Const COL_CTRL_S1   As Integer = 12  ' L - S1 reporting reçu ?
Private Const COL_CTRL_S2   As Integer = 15  ' O - S2 reporting reçu ?
Private Const COL_CTRL_PCA  As Integer = 18  ' R - Bilan PCA reçu ?

' ============================================================
'  UTILITAIRES
' ============================================================
Private Function GetSemestre() As String
    Dim m As Integer: m = Month(Now)
    If m >= 6 And m <= 11 Then
        GetSemestre = "S1"
    Else
        GetSemestre = "S2"
    End If
End Function

Private Function GetDeadlineLabel() As String
    Dim sem As String: sem = GetSemestre()
    If sem = "S1" Then
        GetDeadlineLabel = "30 juin " & Year(Now)
    Else
        GetDeadlineLabel = "30 décembre " & Year(Now)
    End If
End Function

Private Function GetRelanceLabel() As String
    Dim sem As String: sem = GetSemestre()
    Dim deadline As Date
    If sem = "S1" Then
        deadline = DateSerial(Year(Now), 6, 30)
    Else
        deadline = DateSerial(Year(Now), 12, 30)
    End If
    Dim delta As Long: delta = DateDiff("d", Now, deadline)
    Select Case delta
        Case 7:  GetRelanceLabel = "1ère relance (J-7)"
        Case 3:  GetRelanceLabel = "2ème relance (J-3)"
        Case 1:  GetRelanceLabel = "Dernière relance (J-1)"
        Case Else: GetRelanceLabel = "Relance"
    End Select
End Function

Private Function GetRelanceLabelL09() As String
    Dim deadline As Date: deadline = DateSerial(Year(Now), 3, 31)
    Dim delta As Long: delta = DateDiff("d", Now, deadline)
    Select Case delta
        Case 7:  GetRelanceLabelL09 = "1ère relance (J-7)"
        Case 3:  GetRelanceLabelL09 = "2ème relance (J-3)"
        Case 1:  GetRelanceLabelL09 = "Dernière relance (J-1)"
        Case Else: GetRelanceLabelL09 = "Relance"
    End Select
End Function

Private Sub CreerBrouillon(dest As String, sujet As String, corps As String)
    Dim olApp As Object
    Dim mail  As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlook n'est pas disponible. Vérifiez qu'Outlook est installé et ouvert.", vbCritical
        Exit Sub
    End If
    Set mail = olApp.CreateItem(0)
    mail.To      = dest
    mail.Subject = sujet
    mail.HTMLBody = corps
    mail.Display False
End Sub

Private Function SignatureHTML() As String
    SignatureHTML = "<br><br><p><b>" & EXPEDITEUR_NOM & "</b><br>" & _
                   DIRECTION & "<br><i>" & EXPEDITEUR_EMAIL & "</i></p>"
End Function

' ============================================================
'  L1-08 — LANCER CAMPAGNE (mail initial)
' ============================================================
Sub LancerCampagneL108()
    Dim wsDB   As Worksheet: Set wsDB   = ThisWorkbook.Sheets("DATABASE")
    Dim wsCtrl As Worksheet: Set wsCtrl = ThisWorkbook.Sheets("CONTRÔLES LOD1")
    Dim sem    As String: sem = GetSemestre()
    Dim dl     As String: dl  = GetDeadlineLabel()
    Dim compteur As Integer: compteur = 0

    Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, COL_NOM).End(xlUp).Row
    Dim r As Long
    For r = 4 To lastRow
        If UCase(Trim(wsDB.Cells(r, COL_TYPE).Value)) = "PCI" And _
           UCase(Trim(wsDB.Cells(r, COL_INTRAHG).Value)) = "HORS GROUPE" Then
            Dim nom   As String: nom   = Trim(wsDB.Cells(r, COL_NOM).Value)
            Dim rNom  As String: rNom  = Trim(wsDB.Cells(r, COL_RESP_NOM).Value)
            Dim rMail As String: rMail = Trim(wsDB.Cells(r, COL_RESP_MAIL).Value)
            If nom = "" Or rMail = "" Then GoTo NextRow108
            If InStr(rMail, "@") = 0 Then GoTo NextRow108

            Dim sujet As String
            sujet = "[EXT-L1-08 | " & sem & " " & Year(Now) & "] Évaluation semestrielle — " & nom & " — À retourner avant le " & dl

            Dim corps As String
            corps = "<p>Bonjour " & rNom & ",</p>" & _
                    "<p>Dans le cadre du dispositif de contrôle permanent de premier niveau des externalisations " & _
                    "(<b>EXT-L1-08 — Pilotage des événements majeurs</b>), nous vous sollicitons pour la réalisation " & _
                    "de l'évaluation semestrielle concernant la prestation <b>" & nom & "</b> " & _
                    "(" & sem & " " & Year(Now) & ").</p>" & _
                    "<p><b>Date limite de transmission : " & dl & ".</b></p>" & _
                    "<p><b>Ce qui vous est demandé :</b></p><ul>" & _
                    "<li>Confirmer la réception du reporting du prestataire sur la période (indicateurs qualité, incidents, SLA)</li>" & _
                    "<li>Identifier et documenter tout événement majeur ayant conduit à une dégradation du service</li>" & _
                    "<li>Renseigner le statut (Vert / Orange / Rouge) et un commentaire dans la colonne <i>" & sem & "</i> " & _
                    "de l'onglet <b>CONTRÔLES LOD1</b> de la matrice partagée</li></ul>" & _
                    "<p>En l'absence d'événement majeur à signaler, merci de confirmer explicitement en saisissant " & _
                    "<i>Oui / Vert / RAS</i> dans la matrice.</p>" & _
                    "<p>Ce contrôle est obligatoire pour toutes les PCI Hors Groupe et constitue une pièce justificative " & _
                    "exigée dans le cadre de nos obligations EBA/DORA et du Comité de Contrôle Interne (CCI).</p>" & _
                    SignatureHTML()

            CreerBrouillon rMail, sujet, corps
            compteur = compteur + 1
        End If
NextRow108:
    Next r

    If compteur = 0 Then
        MsgBox "Aucun destinataire trouvé (vérifiez les champs Type, Intra/HG et Email dans DATABASE).", vbInformation
    Else
        MsgBox compteur & " brouillon(s) Outlook créé(s) pour la campagne L1-08 " & sem & "." & Chr(10) & _
               "Vérifiez chaque brouillon avant envoi.", vbInformation, "Campagne L1-08 lancée"
    End If
End Sub

' ============================================================
'  L1-08 — RELANCES (non-répondants seulement)
' ============================================================
Sub RelancesL108()
    Dim wsDB   As Worksheet: Set wsDB   = ThisWorkbook.Sheets("DATABASE")
    Dim wsCtrl As Worksheet: Set wsCtrl = ThisWorkbook.Sheets("CONTRÔLES LOD1")
    Dim sem    As String: sem = GetSemestre()
    Dim dl     As String: dl  = GetDeadlineLabel()
    Dim rl     As String: rl  = GetRelanceLabel()
    Dim colStat As Integer
    If sem = "S1" Then colStat = COL_CTRL_S1 Else colStat = COL_CTRL_S2
    Dim compteur As Integer: compteur = 0

    Dim lastRowDB   As Long: lastRowDB   = wsDB.Cells(wsDB.Rows.Count, COL_NOM).End(xlUp).Row
    Dim lastRowCtrl As Long: lastRowCtrl = wsCtrl.Cells(wsCtrl.Rows.Count, COL_CTRL_NOM).End(xlUp).Row

    Dim r As Long
    For r = 4 To lastRowDB
        If UCase(Trim(wsDB.Cells(r, COL_TYPE).Value)) = "PCI" And _
           UCase(Trim(wsDB.Cells(r, COL_INTRAHG).Value)) = "HORS GROUPE" Then
            Dim nom   As String: nom   = Trim(wsDB.Cells(r, COL_NOM).Value)
            Dim rNom  As String: rNom  = Trim(wsDB.Cells(r, COL_RESP_NOM).Value)
            Dim rMail As String: rMail = Trim(wsDB.Cells(r, COL_RESP_MAIL).Value)
            If nom = "" Or rMail = "" Then GoTo NextRowR108
            If InStr(rMail, "@") = 0 Then GoTo NextRowR108

            ' Chercher statut dans CONTRÔLES LOD1
            Dim rc As Long: Dim dejaRendu As Boolean: dejaRendu = False
            For rc = 5 To lastRowCtrl
                If Trim(wsCtrl.Cells(rc, COL_CTRL_NOM).Value) = nom Then
                    If UCase(Trim(wsCtrl.Cells(rc, colStat).Value)) = "OUI" Then
                        dejaRendu = True
                    End If
                    Exit For
                End If
            Next rc

            If Not dejaRendu Then
                Dim sujet As String
                sujet = "[" & rl & " | EXT-L1-08 | " & sem & " " & Year(Now) & "] " & nom & " — Évaluation semestrielle à transmettre avant le " & dl

                Dim corps As String
                corps = "<p>Bonjour " & rNom & ",</p>" & _
                        "<p>Sauf erreur de notre part, nous n'avons pas encore reçu votre évaluation semestrielle " & _
                        "(<b>EXT-L1-08</b>) concernant la prestation <b>" & nom & "</b> pour le " & sem & " " & Year(Now) & ".</p>" & _
                        "<p><b>Date limite : " & dl & ".</b></p>" & _
                        "<p>Merci de renseigner le statut (Vert / Orange / Rouge) et un commentaire dans l'onglet " & _
                        "<b>CONTRÔLES LOD1</b> de la matrice, ou de nous confirmer par retour de mail " & _
                        "l'absence d'événement majeur à signaler (<i>RAS</i>).</p>" & _
                        "<p>Sans retour de votre part avant la date limite, ce contrôle sera enregistré en statut " & _
                        "<b>Rouge</b> et remonté à la Direction des Risques.</p>" & _
                        SignatureHTML()

                CreerBrouillon rMail, sujet, corps
                compteur = compteur + 1
            End If
        End If
NextRowR108:
    Next r

    If compteur = 0 Then
        MsgBox "Tous les prestataires PCI Hors Groupe ont déjà rendu leur évaluation " & sem & ".", vbInformation
    Else
        MsgBox compteur & " brouillon(s) de relance L1-08 créé(s) pour le " & sem & "." & Chr(10) & _
               "Vérifiez chaque brouillon avant envoi.", vbInformation, "Relances L1-08"
    End If
End Sub

' ============================================================
'  L1-09 — LANCER CAMPAGNE (mail initial prestataires)
' ============================================================
Sub LancerCampagneL109()
    Dim wsDB   As Worksheet: Set wsDB   = ThisWorkbook.Sheets("DATABASE")
    Dim annee  As Integer: annee = Year(Now) - 1
    Dim dl     As String: dl = "31 mars " & Year(Now)
    Dim compteur As Integer: compteur = 0

    Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, COL_NOM).End(xlUp).Row
    Dim r As Long
    For r = 4 To lastRow
        If UCase(Trim(wsDB.Cells(r, COL_TYPE).Value)) = "PCI" Then
            Dim nom    As String: nom    = Trim(wsDB.Cells(r, COL_NOM).Value)
            Dim pcaNom As String: pcaNom = Trim(wsDB.Cells(r, COL_PCA_NOM).Value)
            Dim pcaMail As String: pcaMail = Trim(wsDB.Cells(r, COL_PCA_MAIL).Value)
            If nom = "" Or pcaMail = "" Then GoTo NextRow109
            If InStr(pcaMail, "@") = 0 Then GoTo NextRow109

            Dim sujet As String
            sujet = "[EXT-L1-09 | Amundi Immobilier] Demande de bilan PCA " & annee & " — À transmettre avant le " & dl

            Dim corps As String
            corps = "<p>Bonjour " & pcaNom & ",</p>" & _
                    "<p>Dans le cadre de nos obligations contractuelles et réglementaires " & _
                    "(EBA Guidelines on Outsourcing, DORA — contrôle <b>EXT-L1-09</b>), " & _
                    "nous vous sollicitons pour la transmission du <b>bilan annuel de tests PCA</b> " & _
                    "relatif à votre prestation pour <b>Amundi Immobilier</b>, pour l'exercice <b>" & annee & "</b>.</p>" & _
                    "<p><b>Date limite de transmission : " & dl & ".</b></p>" & _
                    "<p><b>Documents attendus :</b></p><ul>" & _
                    "<li>Bilan des tests PCA réalisés en " & annee & " (scénarios couverts, résultats, anomalies éventuelles)</li>" & _
                    "<li>Preuve de réalisation des tests (rapport, procès-verbal ou attestation)</li>" & _
                    "<li>Plan d'action correctif si des anomalies ont été identifiées</li>" & _
                    "<li>Confirmation de l'interaction entre votre dispositif de gestion de crise et celui d'Amundi Immobilier</li></ul>" & _
                    "<p>Ces éléments constituent la piste d'audit requise pour notre dispositif de contrôle interne " & _
                    "et sont susceptibles d'être demandés par nos autorités de supervision (BCE, ACPR/AMF).</p>" & _
                    "<p>Merci de transmettre ces documents à l'adresse : <b>" & EXPEDITEUR_EMAIL & "</b> " & _
                    "en mentionnant en objet la référence de votre contrat avec Amundi Immobilier.</p>" & _
                    SignatureHTML()

            CreerBrouillon pcaMail, sujet, corps
            compteur = compteur + 1
        End If
NextRow109:
    Next r

    If compteur = 0 Then
        MsgBox "Aucun destinataire PCI trouvé (vérifiez les champs Type et Email Contact PCA dans DATABASE).", vbInformation
    Else
        MsgBox compteur & " brouillon(s) Outlook créé(s) pour la campagne L1-09." & Chr(10) & _
               "Vérifiez chaque brouillon avant envoi.", vbInformation, "Campagne L1-09 lancée"
    End If
End Sub

' ============================================================
'  L1-09 — RELANCES (prestataires sans bilan reçu)
' ============================================================
Sub RelancesL109()
    Dim wsDB   As Worksheet: Set wsDB   = ThisWorkbook.Sheets("DATABASE")
    Dim wsCtrl As Worksheet: Set wsCtrl = ThisWorkbook.Sheets("CONTRÔLES LOD1")
    Dim annee  As Integer: annee = Year(Now) - 1
    Dim dl     As String: dl = "31 mars " & Year(Now)
    Dim rl     As String: rl = GetRelanceLabelL09()
    Dim compteur As Integer: compteur = 0

    Dim lastRowDB   As Long: lastRowDB   = wsDB.Cells(wsDB.Rows.Count, COL_NOM).End(xlUp).Row
    Dim lastRowCtrl As Long: lastRowCtrl = wsCtrl.Cells(wsCtrl.Rows.Count, COL_CTRL_NOM).End(xlUp).Row

    Dim r As Long
    For r = 4 To lastRowDB
        If UCase(Trim(wsDB.Cells(r, COL_TYPE).Value)) = "PCI" Then
            Dim nom    As String: nom    = Trim(wsDB.Cells(r, COL_NOM).Value)
            Dim pcaNom  As String: pcaNom  = Trim(wsDB.Cells(r, COL_PCA_NOM).Value)
            Dim pcaMail As String: pcaMail = Trim(wsDB.Cells(r, COL_PCA_MAIL).Value)
            If nom = "" Or pcaMail = "" Then GoTo NextRowR109
            If InStr(pcaMail, "@") = 0 Then GoTo NextRowR109

            Dim rc As Long: Dim dejaRecu As Boolean: dejaRecu = False
            For rc = 5 To lastRowCtrl
                If Trim(wsCtrl.Cells(rc, COL_CTRL_NOM).Value) = nom Then
                    If UCase(Trim(wsCtrl.Cells(rc, COL_CTRL_PCA).Value)) = "OUI" Then
                        dejaRecu = True
                    End If
                    Exit For
                End If
            Next rc

            If Not dejaRecu Then
                Dim sujet As String
                sujet = "[" & rl & " | EXT-L1-09 | Amundi Immobilier] Bilan PCA " & annee & " — Transmission attendue avant le " & dl

                Dim corps As String
                corps = "<p>Bonjour " & pcaNom & ",</p>" & _
                        "<p>Sauf erreur de notre part, nous n'avons pas encore reçu le <b>bilan de tests PCA " & annee & "</b> " & _
                        "relatif à votre prestation pour Amundi Immobilier (<b>EXT-L1-09</b>).</p>" & _
                        "<p><b>Date limite : " & dl & ".</b></p>" & _
                        "<p>Sans réception de ces documents avant la date limite, nous serons contraints d'enregistrer " & _
                        "ce contrôle en statut <b>Rouge</b> et d'en informer notre Direction des Risques, " & _
                        "conformément à nos obligations réglementaires (DORA, EBA).</p>" & _
                        "<p>Pour rappel, les documents attendus sont :" & _
                        "<ul><li>Bilan des tests PCA " & annee & " (scénarios, résultats)</li>" & _
                        "<li>Preuve de réalisation</li>" & _
                        "<li>Plan d'action correctif si anomalies</li></ul></p>" & _
                        "<p>Merci de transmettre ces éléments à : <b>" & EXPEDITEUR_EMAIL & "</b></p>" & _
                        SignatureHTML()

                CreerBrouillon pcaMail, sujet, corps
                compteur = compteur + 1
            End If
        End If
NextRowR109:
    Next r

    If compteur = 0 Then
        MsgBox "Tous les bilans PCA ont été reçus.", vbInformation
    Else
        MsgBox compteur & " brouillon(s) de relance L1-09 créé(s)." & Chr(10) & _
               "Vérifiez chaque brouillon avant envoi.", vbInformation, "Relances L1-09"
    End If
End Sub
