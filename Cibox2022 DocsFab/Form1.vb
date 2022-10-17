Imports Microsoft.Office.Interop
Imports iTextSharp.text
Imports System.Drawing
Imports System.Threading
Imports System.Text.RegularExpressions
Imports Rectangle = iTextSharp.text.Rectangle
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser

Public Class Form1
    Dim source As String
    Dim resultat As Integer
    Dim Names_Files(400) As String
    Dim Names_Files_plans(400) As String
    Dim Names_Files_plans_prises As String = ""
    Dim _strCustomerCSVPath As String
    Dim _Ids_OUV(300) As Long, _Ids_DOR(300) As Long, _Ids_POT(300) As Long, _Ids_MEN(300) As Long, _Ids_FIX(300) As Long
    Dim j__OUV As Integer, j__DOR As Integer, j__POT As Integer, j__MEN As Integer, j__FIX As Integer = 0
    Dim str__OUV As String, str__DOR As String, str__POT As String, str__MEN As String, str__FIX As String = ","
    Dim doc As New Document()
    Dim copier As New PdfCopy(doc, New FileStream("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF", FileMode.Create))
    Dim doc_non_trie As New Document()
    Dim copier_non_trie As New PdfCopy(doc_non_trie, New FileStream("C:\Traitement DiffusionDocsFab\DOSSIER_NON_TRIE.PDF", FileMode.Create))
    Dim _comp_ligne(15) As Integer
    Dim _Entete_service(15) As Integer
    Dim _Condition_Cib(15) As Boolean
    Dim _comp = 0
    Dim p_plan As Integer = 0
    Dim pp_plan As Integer = 0
    Dim repitition As String = ""
    Dim _nombre_plans As Integer = 0
    Dim _Fichier_Log As String = "C:\FichierLog_DiffusionDocsFab_V13_10_22.txt"
    Dim condition_Exp As Boolean = False
    Dim year As String = ""
    Dim path_name As String = "C:\Name.txt"
    Dim _cmpt_total As Integer = 0
    Dim _cmpt_SOUDAGE(2) As Integer
    Dim _cmpt_MAGASIN(2) As Integer
    Dim _cmpt_PEINTURE(2) As Integer
    Dim _cmpt_ALU(2) As Integer
    Dim _cmpt_PLANS(2) As Integer
    Dim _cmpt_PLANS_S60(2) As Integer
    Dim _cmpt_PLANS_TUBE(2) As Integer
    Dim _cmpt_PLANS_S00(2) As Integer
    Dim path_export As String
    Dim LISTE_comp As String

    Dim _Pages_Etat(40) As Integer
    Dim _indice_Pages_Etat As Integer = 0
    Dim file_path_trace As String

    Dim listecomp(100) As String
    Dim inlistecomp As Integer = 0

    Dim XLS1 As Excel.Application

    Dim listecompSou(100) As String
    Dim inlistecompSou As Integer = 0

    Dim pathExportVault As String

    Dim Classeur1 As Microsoft.Office.Interop.Excel._Workbook
    Dim destExcelFileRoot_FeuilleComposant As String = "C:\Traitement DiffusionDocsFab\FeuillesComposants.xls"

    Private Function TestOpenDirectory() As String
        Dim str As String
        If ComboBox1.SelectedItem.ToString() <> "" Then
            year = "C" + ComboBox1.SelectedItem.ToString()
        End If

        Select Case Len(TextBox1.Text)
            Case 1
                TextBox1.Text = year + "-" + "000" + TextBox1.Text.ToString
            Case 2
                TextBox1.Text = year + "-" + "00" + TextBox1.Text.ToString
            Case 3
                TextBox1.Text = year + "-" + "0" + TextBox1.Text.ToString
                'MsgBox(TextBox1.Text)
            Case 4
                TextBox1.Text = year + "-" + TextBox1.Text.ToString

        End Select

        If System.IO.Directory.Exists("\\srvdc\Autocad\PDF-FAB\" + TextBox1.Text) Then

            str = "\\srvdc\Autocad\PDF-FAB\" + TextBox1.Text


        Else


            MsgBox("Erreur de saisie !")
            str = ""
            TextBox1.Text = ""
            Label1.Text = "Veuillez renseigner le numéro du dossier !"

        End If
        Return str

    End Function

    'Control du bouton par Entrer
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.AcceptButton = Button1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'resultat = Inistialisation()
        Dim _Condition_Meneau As Boolean = False

        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False And RadioButton8.Checked = False Then

            MsgBox("Veuillez choisir une finition !")
        Else
            Label1.Text = "Traitement en cours !"
            If TextBox2.Text = "" Then

                Dim p_name As New System.IO.StreamReader(path_name)
                While p_name.Peek() >= 0
                    Dim name As String = p_name.ReadLine()
                    If name <> "" Then
                        TextBox2.Text = name

                    End If
                End While
                p_name.Close()
            Else
                My.Computer.FileSystem.WriteAllText(path_name, TextBox2.Text + vbCrLf, True)

            End If

            source = TestOpenDirectory()

            'Run Code Source Jack''''''''''''''''''''''''''''''''''''''
            Try

                Dim p As Process = Process.Start("\\srvdc\Bureau_Etudes\instal_BE\TRAITEDXF\setup.exe")
                p.WaitForExit()

            Catch ex As Exception

                MsgBox(ex.Message)

            End Try


            _strCustomerCSVPath = "C:\Traitement DiffusionDocsFab\" + TextBox1.Text + ".csv"
            If My.Computer.FileSystem.FileExists(_strCustomerCSVPath) Then
                My.Computer.FileSystem.DeleteFile(_strCustomerCSVPath)
            End If
            If My.Computer.FileSystem.FileExists(_Fichier_Log) Then
                My.Computer.FileSystem.DeleteFile(_Fichier_Log)
            End If
            If System.IO.Directory.Exists("C:\Traitement DiffusionDocsFab") Then
                For Each files As String In System.IO.Directory.GetFiles("C:\Traitement DiffusionDocsFab")
                    If files.Contains("DOSSIER") <> True Then
                        System.IO.File.Delete(files)
                    End If
                Next
            End If

            If source <> "" Then

                If RadioButton7.Checked = True Then

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("Export") Then
                            pathExportVault = files
                        End If
                    Next
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************************************************** Début traitement :) ***********************************************************" + vbCrLf, True)
                    doc_non_trie.Open()
                    resultat = _CibSlide_Premiertrie()
                    resultat = _Cibslide_Deuxiemetrie()
                    resultat = _Cibslide_Cinquièmetrie()


                    ''''''''''''''''''''''''''''''''''''S00 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    doc.Open()

                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S00*********************" + vbCrLf, True)
                    resultat = GetDataFromCsv("Accessoires", "Page_non_prise", "")

                    resultat = GetDataFromCsv("Liste_Encadrement", "", "")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ENCADREMENTS") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next



                    'Plan Tube /Doc Soudage + BF
                    resultat = GetDataFromCsv("Plan_TUBE", "ATELIER", "")


                    ' Etat pièces

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\00.pdf", "", 1)


                    resultat = GetDataFromCsv("Etat_Pieces", "OUVRANT", "S00")
                    resultat = GetDataFromCsv("Etat_Pieces", "DORMANT", "S00")

                    resultat = GetDataFromCsv("Etat_Pieces", "U134", "S00")
                    resultat = GetDataFromCsv("Etat_Pieces", "TUBE", "S00")
                    resultat = GetDataFromCsv("Etat_Pieces", "POTEAU", "S00")
                    resultat = GetDataFromCsv("Etat_Pieces", "Page_non_prise", "S00")


                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S00")

                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "Page_non_prise", "S00")





                    inlistecomp = 0
                    While listecomp(inlistecomp) <> Nothing
                        resultat = AddPages("C:\Traitement DiffusionDocsFab\Feuille " + listecomp(inlistecomp) + ".pdf", "", 1)

                        If listecomp(inlistecomp).StartsWith("MENEAU") Then
                            _Condition_Meneau = True
                            resultat = GetDataFromCsv("Etat_Pieces", "MENEAU", "S00")
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S00")
                        End If

                        For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                            If Split(files, "\")(7).StartsWith("SOUDAGE " + listecomp(inlistecomp)) Then

                                resultat = AddPages(files, "", 1)


                                listecompSou(inlistecompSou) = Split(Split(files, "\")(7), "SOUDAGE ")(1)
                                inlistecompSou = inlistecompSou + 1
                            End If
                        Next

                        resultat = GetDataFromCsv("Plan_S00", listecomp(inlistecomp), "S00")

                        inlistecomp = inlistecomp + 1
                    End While



                    'SOUDAGE RESTANTS

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")

                        If Split(files, "\")(7).StartsWith("SOUDAGE") Then
                            If listecompSou.Contains(Split(Split(files, "\")(7), "SOUDAGE ")(1)) <> True Then
                                MsgBox(Split(Split(files, "\")(7), "SOUDAGE ")(1))
                                resultat = AddPages(files, "", 1)

                            End If
                        End If
                    Next


                    '''''''''''''''''''''''''Fixe'''''''''''''''''''''''''

                    resultat = GetDataFromCsv("Etat_Pieces", "FIXE", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S00")
                    resultat = GetDataFromCsv("Plan_S00", "FIXE", "S00")
                    resultat = GetDataFromCsv("Plan_S00", "Non renseigne", "S00")

                    If _Condition_Meneau = False Then
                        resultat = GetDataFromCsv("Etat_Pieces", "MENEAU", "S00")
                        resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S00")
                    End If

                    'resultat = _Cibslide_Cinquièmetrie()
                    ''''''''''''''''''''''''''''''''''''S10 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S10*********************" + vbCrLf, True)
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 1)
                    'S10 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S10")

                    resultat = GetDataFromCsv("Liste_Encadrement", "", "")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ENCADREMENTS") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next

                    resultat = GetDataFromCsv("Etat_Pieces", "OUVRANT", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "DORMANT", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "FIXE", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "MENEAU", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "U134", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "TUBE", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "POTEAU", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "Page_non_prise", "S10")

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("PEINTURE") Then
                            resultat = AddPages(files, "", 1)

                        End If
                    Next
                    ''''''''''''''''''''''''''''''''''''S20 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 2)
                    'S20 Fait
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S20*********************" + vbCrLf, True)
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "QUINCAILLERIE", "S20")
                    resultat = GetDataFromCsv("Accessoires", "Expedition_Finition", "S20")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("EXPEDITION") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next


                    ''''''''''''''''''''''''''''''''''''S30 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 3)
                    'S30 Fait
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S30*********************" + vbCrLf, True)
                    resultat = GetDataFromCsv("BT ALUMINIUM", "", "S30")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "ALU", "S30")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S30")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ALU") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next
                    resultat = GetDataFromCsv("ALU", "ALU", "S30")
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S30")

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ENCADREMENTS") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next

                    ''''''''''''''''''''''''''''''''''''S40 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    '''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 4)
                    'S40 Fait
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S40*********************" + vbCrLf, True)
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S40")

                    ''''''''''''''''''''''''''''''''''''S50 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 5)
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S50*********************" + vbCrLf, True)
                    'S50 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S50")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S50")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S50")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S50")


                    ''''''''''''''''''''''''''''''''''''S60 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 6)
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S60*********************" + vbCrLf, True)
                    'S60 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S60")

                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S60")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ENCADREMENTS") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next

                    resultat = GetDataFromCsv("Accessoires", "Expedition_Finition", "S60")
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("Expédition") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next
                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("MONTAGE") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next
                    resultat = GetDataFromCsv("Plan_S60", "MONTAGE", "S60")


                    ''''''''''''''''''''''''''''''''''''S70 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 7)
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S70*********************" + vbCrLf, True)
                    'S70 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "QUINCAILLERIE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "ALU", "S70")


                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ALU") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next


                    ''''''''''''''''''''''''''''''''''''S90 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 9)
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S90*********************" + vbCrLf, True)
                    'S90 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S90")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S90")

                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S90")

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("MONTAGE VITRAGE") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next


                    ''''''''''''''''''''''''''''''''''''S100 CibSlide''''''''''''''''''''''''''''''''''''''''''''
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 10)
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************S100********************" + vbCrLf, True)
                    'S100 Fait
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S100")

                    For Each files As String In System.IO.Directory.GetFiles(source + "\PDFS")
                        If Split(files, "\")(7).StartsWith("ENCADREMENTS") Then
                            resultat = AddPages(files, "", 1)
                        End If
                    Next

                    doc.Close()
                    copier.Close()
                    copier_non_trie.Close()
                    doc_non_trie.Close()


                    If RadioButton1.Checked = True Or (RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False And RadioButton8.Checked = False) Then

                        If My.Computer.FileSystem.FileExists("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF") Then
                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN_TRIE\" + TextBox1.Text + "_Pôles.PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_NON_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN\" + TextBox1.Text + ".PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                        End If

                    Else
                        resultat = WaterMark("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF")

                        If My.Computer.FileSystem.FileExists("C:\Traitement DiffusionDocsFab\" + TextBox1.Text + "_Pôles.PDF") Then
                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\" + TextBox1.Text + "_Pôles.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN_TRIE\" + TextBox1.Text + "_Pôles.PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_NON_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN\" + TextBox1.Text + ".PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                        End If
                    End If

                    Label1.Text = " Le dossier trié se trouve dans : CatalnomPLAN_TRIE"

                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************************************************** Fin traitement :) ***********************************************************" + vbCrLf, True)

                Else

                    MsgBox("Traitement en cours !")
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************************************************** Début traitement :) ***********************************************************" + vbCrLf, True)
                    resultat = AddDatatoCSVfile("_Type_Document", "_Composant", "_page", "_File_path")
                    resultat = Premiertrie()
                    resultat = Deuxiemetrie()
                    resultat = AddDatatoCSVfile("IdsPlans OUV : ", str__OUV, "", "")
                    resultat = AddDatatoCSVfile("IdsPlans DOR : ", str__DOR, "", "")
                    resultat = AddDatatoCSVfile("IdsPlans POT : ", str__POT, "", "")
                    resultat = AddDatatoCSVfile("IdsPlans MEN : ", str__MEN, "", "")
                    resultat = AddDatatoCSVfile("IdsPlans FIX : ", str__FIX, "", "")

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 'Add Data''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    resultat = GetDataFromCsv("Accessoires", "Page_non_prise", "")
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "")
                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\00.pdf", "", 1)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S00''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    resultat = GetDataFromCsv("Etat_Pieces", "TUBE", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure TUBE", "S00")

                    '----------Plan S00 TUBE
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans TUBE/BUTEE-------------" + vbCrLf, True)
                    resultat = Correction_plans("TUBE", "S00")
                    resultat = Correction_plans("BUTEE", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure BUTEE", "S00")

                    If _Condition_Cib(3) = True Then
                        resultat = GetDataFromCsv("Etat_Pieces", "OUVRANT", "S00")
                        resultat = GetDataFromCsv("Etat_Pieces", "DORMANT", "S00")
                        resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S00")
                        resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S00")
                        resultat = GetDataFromCsv("Accessoires", "Soudure OUVRANT", "S00")
                        resultat = GetDataFromCsv("Accessoires", "Soudure DORMANT", "S00")

                        '----------Plan S00 OUVRANT
                        p_plan = 0
                        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans OUVRANT-------------" + vbCrLf, True)
                        While _Ids_OUV(p_plan) <> 0

                            If System.IO.File.Exists(source + "\" + _Ids_OUV(p_plan).ToString + ".pdf") Then

                                'resultat = AddPages(source + _Ids_MEN(p_plan).ToString, "S00", 1)
                                _nombre_plans = _nombre_plans + 1
                                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_OUV(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                                resultat = AddPages(source + "\" + _Ids_OUV(p_plan).ToString + ".pdf", "S00", 1)
                                Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_OUV(p_plan).ToString + ".pdf"
                            Else

                                If Len(_Ids_OUV(p_plan).ToString) >= 8 Then
                                    Dim testString As String = _Ids_OUV(p_plan).ToString
                                    Dim subString As String = testString.Substring(6)
                                    resultat = AddPlans(subString, "S00", "OUVRANT")
                                    resultat = AddPlans(subString, "S00", "BLOC")
                                End If

                            End If
                            p_plan = p_plan + 1
                        End While
                        resultat = Correction_plans("OUVRANT", "S00")
                        resultat = Correction_plans("BLOC", "S00")

                        '----------Plan S00 DORMANT
                        p_plan = 0
                        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans DORMANT-------------" + vbCrLf, True)
                        While _Ids_DOR(p_plan) <> 0
                            If System.IO.File.Exists(source + "\" + _Ids_DOR(p_plan).ToString + ".pdf") Then

                                'resultat = AddPages(source + _Ids_MEN(p_plan).ToString, "S00", 1)
                                _nombre_plans = _nombre_plans + 1
                                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_DOR(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                                resultat = AddPages(source + "\" + _Ids_DOR(p_plan).ToString + ".pdf", "S00", 1)
                                Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_DOR(p_plan).ToString + ".pdf"
                            Else

                                If Len(_Ids_DOR(p_plan).ToString) >= 8 Then
                                    Dim testString As String = _Ids_DOR(p_plan).ToString
                                    Dim subString As String = testString.Substring(6)
                                    resultat = AddPlans(subString, "S00", "DORMANT")
                                End If

                            End If
                            p_plan = p_plan + 1


                        End While
                        resultat = Correction_plans("DORMANT", "S00")
                        resultat = Correction_plans("SEUIL", "S00")


                    Else
                        resultat = GetDataFromCsv("Etat_Pieces", "OUVRANT", "S00")
                        resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S00")
                        resultat = GetDataFromCsv("Accessoires", "Soudure OUVRANT", "S00")
                        '----------Plan S00 OUVRANT
                        p_plan = 0
                        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans OUVRANT-------------" + vbCrLf, True)
                        While _Ids_OUV(p_plan) <> 0

                            If System.IO.File.Exists(source + "\" + _Ids_OUV(p_plan).ToString + ".pdf") Then

                                'resultat = AddPages(source + _Ids_MEN(p_plan).ToString, "S00", 1)
                                _nombre_plans = _nombre_plans + 1
                                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_OUV(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                                resultat = AddPages(source + "\" + _Ids_OUV(p_plan).ToString + ".pdf", "S00", 1)
                                Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_OUV(p_plan).ToString + ".pdf"
                            Else

                                If Len(_Ids_OUV(p_plan).ToString) >= 8 Then
                                    Dim testString As String = _Ids_OUV(p_plan).ToString
                                    Dim subString As String = testString.Substring(6)
                                    resultat = AddPlans(subString, "S00", "OUVRANT")
                                    resultat = AddPlans(subString, "S00", "BLOC")
                                End If

                            End If
                            p_plan = p_plan + 1
                        End While
                        resultat = Correction_plans("OUVRANT", "S00")
                        resultat = Correction_plans("BLOC", "S00")


                        resultat = GetDataFromCsv("Etat_Pieces", "DORMANT", "S00")
                        resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S00")
                        resultat = GetDataFromCsv("Accessoires", "Soudure DORMANT", "S00")
                        '----------Plan S00 DORMANT
                        p_plan = 0
                        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans DORMANT-------------" + vbCrLf, True)
                        While _Ids_DOR(p_plan) <> 0
                            'MsgBox(_Ids_DOR(p_plan))

                            If System.IO.File.Exists(source + "\" + _Ids_DOR(p_plan).ToString + ".pdf") Then

                                'resultat = AddPages(source + _Ids_MEN(p_plan).ToString, "S00", 1)
                                _nombre_plans = _nombre_plans + 1
                                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_DOR(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                                resultat = AddPages(source + "\" + _Ids_DOR(p_plan).ToString + ".pdf", "S00", 1)
                                Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_DOR(p_plan).ToString + ".pdf"
                            Else

                                If Len(_Ids_DOR(p_plan).ToString) >= 8 Then
                                    Dim testString As String = _Ids_DOR(p_plan).ToString
                                    Dim subString As String = testString.Substring(6)
                                    'MsgBox(subString)
                                    resultat = AddPlans(subString, "S00", "DORMANT")
                                End If

                            End If

                            p_plan = p_plan + 1
                        End While
                        resultat = Correction_plans("DORMANT", "S00")
                        resultat = Correction_plans("SEUIL", "S00")

                    End If
                    resultat = GetDataFromCsv("Etat_Pieces", "FIXE", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure FIXE", "S00")
                    '----------Plan S00 FIXE
                    p_plan = 0
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans FIXE-------------" + vbCrLf, True)
                    While _Ids_FIX(p_plan) <> 0
                        If System.IO.File.Exists(source + "\" + _Ids_FIX(p_plan).ToString + ".pdf") Then

                            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_FIX(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                            resultat = AddPages(source + "\" + _Ids_FIX(p_plan).ToString + ".pdf", "S00", 1)
                            Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_FIX(p_plan).ToString + ".pdf"
                            _nombre_plans = _nombre_plans + 1
                        Else

                            If Len(_Ids_FIX(p_plan).ToString) >= 8 Then
                                Dim testString As String = _Ids_FIX(p_plan).ToString
                                Dim subString As String = testString.Substring(6)
                                resultat = AddPlans(subString, "S00", "FIXE")
                            End If

                        End If
                        p_plan = p_plan + 1

                    End While
                    resultat = Correction_plans("FIXE", "S00")

                    resultat = GetDataFromCsv("Etat_Pieces", "MENEAU", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure MENEAU", "S00")
                    '----------Plan S00 MENEAU
                    p_plan = 0
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans MENEAU-------------" + vbCrLf, True)
                    While _Ids_MEN(p_plan) <> 0
                        If System.IO.File.Exists(source + "\" + _Ids_MEN(p_plan).ToString + ".pdf") Then


                            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_MEN(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                            _nombre_plans = _nombre_plans + 1
                            resultat = AddPages(source + "\" + _Ids_MEN(p_plan).ToString + ".pdf", "S00", 1)
                            Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_MEN(p_plan).ToString + ".pdf"
                        Else


                            If Len(_Ids_MEN(p_plan).ToString) >= 8 Then
                                Dim testString As String = _Ids_MEN(p_plan).ToString
                                Dim subString As String = testString.Substring(6)
                                resultat = AddPlans(subString, "S00", "MENEAU")
                                resultat = AddPlans(subString, "S00", "MN")
                            End If

                        End If
                        p_plan = p_plan + 1
                    End While
                    resultat = Correction_plans("MENEAU", "S00")
                    resultat = Correction_plans("MN", "S00")

                    resultat = GetDataFromCsv("Etat_Pieces", "U134", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure U134", "S00")
                    '----------Plan S00 U134
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans U134-------------" + vbCrLf, True)
                    resultat = Correction_plans("U134", "S00")

                    '----------Plan S00 POTEAU
                    resultat = GetDataFromCsv("Etat_Pieces", "POTEAU", "S00")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S00")
                    resultat = GetDataFromCsv("Accessoires", "Soudure POTEAU", "S00")

                    p_plan = 0
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans POTEAU-------------" + vbCrLf, True)
                    While _Ids_POT(p_plan) <> 0
                        If System.IO.File.Exists(source + "\" + _Ids_POT(p_plan).ToString + ".pdf") Then

                            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + source + "\" + _Ids_POT(p_plan).ToString + ".pdf" + " a été pris" + vbCrLf, True)
                            resultat = AddPages(source + "\" + _Ids_POT(p_plan).ToString + ".pdf", "S00", 1)
                            Names_Files_plans_prises = Names_Files_plans_prises + " , " + source + "\" + _Ids_POT(p_plan).ToString + ".pdf"
                            _nombre_plans = _nombre_plans + 1
                        Else

                            If Len(_Ids_POT(p_plan).ToString) >= 8 Then
                                Dim testString As String = _Ids_POT(p_plan).ToString
                                Dim subString As String = testString.Substring(6)
                                resultat = AddPlans(subString, "S00", "POTEAU")
                            End If

                        End If
                        p_plan = p_plan + 1


                    End While
                    resultat = Correction_plans("POTEAU", "S00")


                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 1)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S10''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S10 Fait

                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S10")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S10")
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S10")
                    resultat = GetDataFromCsv("Etat_Pieces", "OUVRANT", "S10")
                    resultat = GetDataFromCsv("Accessoires", "Peinture", "S10")

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 2)

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S20''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S20 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "QUINCAILLERIE", "S20")
                    'resultat = GetDataFromCsv("Accessoires_Exp", "", "S20")
                    resultat = GetDataFromCsv("Accessoires", "Expedition_Finition", "S20")

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 3)

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S30''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S30 Fait
                    resultat = GetDataFromCsv("BT ALUMINIUM", "", "S30")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "ALU", "S30")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S30")

                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans ALU-------------" + vbCrLf, True)
                    '----------Plans ALU
                    p_plan = 0
                    resultat = GetDataFromCsv("Vide", "", "S30")
                    While Names_Files_plans(p_plan) <> ""
                        If Names_Files_plans_prises.Contains(Names_Files_plans(p_plan)) <> True And Names_Files_plans(p_plan).Contains("PAC02601") <> True Then

                            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + Names_Files_plans(p_plan) + " a été pris" + vbCrLf, True)
                            AddPages(Names_Files_plans(p_plan), "S30", 1)
                            Names_Files_plans_prises = Names_Files_plans_prises + "," + Names_Files_plans(p_plan)

                        End If
                        p_plan = p_plan + 1
                    End While
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S30")



                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------Plans Restants-------------" + vbCrLf, True)
                    p_plan = 0
                    While Names_Files_plans(p_plan) <> ""
                        If Names_Files_plans_prises.Contains(Names_Files_plans(p_plan)) <> True Then

                            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + Names_Files_plans(p_plan) + " n'a pas été pris" + vbCrLf, True)
                        End If
                        p_plan = p_plan + 1
                    End While


                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 4)


                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S40''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S40 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S40")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S40")

                    If _Condition_Cib(0) = True Or _Condition_Cib(1) = True Or _Condition_Cib(2) = True Then

                        resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 5)
                        'S50 Fait

                        'case Collé
                        If _Condition_Cib(0) = True Then
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S50")
                        End If

                        'Case Cibisol
                        If _Condition_Cib(1) = True Then
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S50")
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S50")
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S50")
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S50")
                        End If
                        'case Cibfeu
                        If _Condition_Cib(2) = True Then
                            resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S50")
                        End If
                    End If

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 6)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S60''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S60 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S60")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S60")
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S60")
                    'resultat = GetDataFromCsv("Accessoires_Exp", "", "S60")
                    resultat = GetDataFromCsv("Accessoires", "Expedition_Finition", "S60")

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 7)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S70''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S70 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "DORMANT", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "MENEAU", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "U134", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "TUBE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "POTEAU", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "CJ", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "QUINCAILLERIE", "S70")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "ALU", "S70")


                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 9)

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S90''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S90 Fait
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "OUVRANT", "S90")
                    resultat = GetDataFromCsv("Bordereau_Fabrication", "FIXE", "S90")
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S90")

                    resultat = AddPages("\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\Diffusion docs.pdf", "", 10)

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''S100''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'S100 Fait
                    resultat = GetDataFromCsv("Liste_Encadrement", "", "S100")

                    doc.Close()
                    copier.Close()
                    copier_non_trie.Close()
                    doc_non_trie.Close()

                    'La copie des dossier dans les bons chemins

                    If RadioButton1.Checked = True Or (RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False And RadioButton8.Checked = False) Then

                        If My.Computer.FileSystem.FileExists("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF") Then
                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN_TRIE\" + TextBox1.Text + "_Pôles.PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_NON_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN\" + TextBox1.Text + ".PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                        End If

                    Else
                        resultat = WaterMark("C:\Traitement DiffusionDocsFab\DOSSIER_TRIE.PDF")

                        If My.Computer.FileSystem.FileExists("C:\Traitement DiffusionDocsFab\" + TextBox1.Text + "_Pôles.PDF") Then
                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\" + TextBox1.Text + "_Pôles.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN_TRIE\" + TextBox1.Text + "_Pôles.PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

                            My.Computer.FileSystem.CopyFile("C:\Traitement DiffusionDocsFab\DOSSIER_NON_TRIE.PDF", "\\srvdc\Documents\Plan\CatalnomPLAN\" + TextBox1.Text + ".PDF",
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                        End If
                    End If

                    Label1.Text = " Le dossier trié se trouve dans : CatalnomPLAN_TRIE"
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "*********************************************************** Fin traitement :) ***********************************************************" + vbCrLf, True)
                End If

            Else

                TextBox1.Text = ""
                    Label1.Text = "Veuillez renseigner le numéro du dossier !"

            End If


            MsgBox("Fin de traitement !")
            CloseMe()

        End If

    End Sub

    'Traitement
    Private Function Premiertrie() As Integer
        Dim i As Integer = 0
        Dim j As Integer = 0

        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------- 1/Premier Trie -------------" + vbCrLf, True)
        For Each files As String In System.IO.Directory.GetFiles(source)

            If Split(files, source)(1).Contains(Split(TextBox1.Text, "-")(1)) Or Split(files, source)(1).Contains(Split(TextBox1.Text, "-")(1).Substring(1, 3)) = True Or Split(files, source)(1).Contains(Split(TextBox1.Text, "-")(1).Substring(2, 2)) = True Or Split(files, source)(1).Contains(Split(TextBox1.Text, "-")(1).Substring(3, 1)) = True And files <> "" And files.Contains(".pdf") And files.Contains("DOC") <> True And files.Contains("Doc") <> True And files.Contains("doc") <> True Then
                resultat = Fusion(files)
                Names_Files_plans(j) = files.ToString
                j = j + 1
            ElseIf files <> "" And files.Contains(".pdf") Then
                Names_Files(i) = files.ToString

                i = i + 1
            End If

        Next

        Return 1
    End Function

    Private Function Deuxiemetrie() As Integer

        Dim i = 0
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------- 2/Deuxième trie -------------" + vbCrLf, True)
        While Names_Files(i) <> ""
            'MsgBox(Names_Files(i))
            resultat = Fusion(Names_Files(i))
            resultat = TroisièmeTrie(Names_Files(i))
            i = i + 1
        End While

        Return 1
    End Function

    Private Function TroisièmeTrie(path_file As String) As Integer
        Dim condition As Boolean = False
        Dim pr As New PdfReader(path_file)
        Dim numberOfPages As Integer = pr.NumberOfPages
        Dim str As String = ""
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------- 3/Troisième trie, path : " + path_file + " -------------" + vbCrLf, True)

        For i = 1 To numberOfPages

            str = Scan(path_file, i)
            'MsgBox(str)

            'case cib
            If (str.Contains("COLLEE") = True Or str.Contains("C O L L E E") = True Or str.Contains("C O L L E") = True Or str.Contains("COLLE") = True) And str.Contains("B o r d e r e a u P r é p a r a t i o n A c c e s s o i r e s") <> True And str.Contains("Accessoires") <> True Then
                _Condition_Cib(0) = True

            End If

            'Case Cibisol
            If str.Contains("CIB ISOL") = True Or str.Contains("I S O L") = True Then
                _Condition_Cib(1) = True
            End If

            'case Cibfeu
            If str.Contains("CIB FEU") = True Or str.Contains("F E U") Then
                _Condition_Cib(2) = True
            End If
            'case Cave
            If str.Contains("CAVE") = True Then
                _Condition_Cib(3) = True
            End If

            If str.Contains("ENCADREMENTS") Then
                resultat = AddDatatoCSVfile("Liste_Encadrement", "Vide", i, path_file)
            ElseIf str.Contains("BT ALUMINIUM") Then
                resultat = AddDatatoCSVfile("BT ALUMINIUM", "Vide", i, path_file)

            ElseIf str.Contains("Etat") Or str.Contains("Pièces") Then
                resultat = Quatrièmetrie("Etat_Pieces", path_file, str, i)

            ElseIf str.Contains("Accessoires") Or str.Contains("A c c e s s o i r e s") Or str.Contains("P r é p a r a t i o n") Then

                resultat = cinquièmetrie("Accessoires", path_file, str, i)

            ElseIf str.Contains("Bodereau") Or str.Contains("Fabrication") Or str.Contains("F a b r i c a t i o n") Then
                resultat = Quatrièmetrie("Bordereau_Fabrication", path_file, str, i)


            ElseIf str.Contains("Perçage") Then
                resultat = Quatrièmetrie("Plan_ALU", path_file, str, i)
            Else
                resultat = AddDatatoCSVfile("Vide", "Pages_non_prise", i, path_file)
            End If

        Next
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-------------- 3/Fin troisième trie, path : " + path_file + " -------------" + vbCrLf, True)

        Return 1

    End Function

    Private Function AddDatatoCSVfile(_Document As String, _Composant As String, _page As String, _File_path As String) As Integer
        _comp = _comp + 1
        Select Case _Document

            Case "Etat_Pieces"
                If _comp_ligne(0) = 0 Then

                    _comp_ligne(0) = _comp
                    _comp_ligne(1) = _comp
                Else
                    _comp_ligne(1) = _comp
                End If

            Case "Bordereau_Fabrication"
                If _comp_ligne(2) = 0 Then

                    _comp_ligne(2) = _comp
                    _comp_ligne(3) = _comp
                Else
                    _comp_ligne(3) = _comp
                End If

            Case "Accessoires_Exp"
                If _comp_ligne(4) = 0 Then

                    _comp_ligne(4) = _comp
                    _comp_ligne(5) = _comp
                Else
                    _comp_ligne(5) = _comp
                End If

            Case "Accessoires"
                If _comp_ligne(6) = 0 Then

                    _comp_ligne(6) = _comp
                    _comp_ligne(7) = _comp
                Else
                    _comp_ligne(7) = _comp
                End If

            Case "BT ALUMINIUM"
                If _comp_ligne(8) = 0 Then

                    _comp_ligne(8) = _comp
                    _comp_ligne(9) = _comp
                Else
                    _comp_ligne(9) = _comp
                End If

            Case "Liste_Encadrement"
                If _comp_ligne(10) = 0 Then

                    _comp_ligne(10) = _comp
                    _comp_ligne(11) = _comp
                Else
                    _comp_ligne(11) = _comp
                End If

            Case "Vide"
                If _comp_ligne(12) = 0 Then

                    _comp_ligne(12) = _comp
                    _comp_ligne(13) = _comp
                Else
                    _comp_ligne(13) = _comp
                End If


        End Select

        Try
            Dim objWriter As System.IO.StreamWriter = System.IO.File.AppendText(_strCustomerCSVPath)
            If System.IO.File.Exists(_strCustomerCSVPath) Then
                objWriter.Write(_Document & ",")
                objWriter.Write(_Composant & ",")
                objWriter.Write(_page & ",")
                objWriter.Write(_File_path & ",")
                objWriter.Write(Environment.NewLine)
            End If
            objWriter.Close()


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return 1

    End Function
    Private Function Scan(path As String, i As Integer) As String

        Dim pr As New PdfReader(path)
        Dim its As ITextExtractionStrategy
        its = New LocationTextExtractionStrategy
        Dim s As String
        s = PdfTextExtractor.GetTextFromPage(pr, i, its)

        Return s
    End Function


    'Quitter
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim iExit As DialogResult

        iExit = MessageBox.Show("Confirmez que vous voulez quitter le système ? ", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If iExit = DialogResult.Yes Then

            Me.Close()
        End If


    End Sub

    Private Function Quatrièmetrie(_type_Document As String, path As String, texte As String, i As Integer) As Integer
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->Sous trie, Page : " + i.ToString + ", " + _type_Document + vbCrLf, True)

        If ((texte.Contains("OUVRANT") Or texte.Contains("OUV") Or texte.Contains("Ouvrant") Or texte.Contains("CORNIERE")) And texte.Contains("MENEAU") <> True And texte.Contains("FIXE") <> True And texte.Contains("IMPOSTE") <> True And texte.Contains("MONTANT") <> True And texte.Contains("MN") <> True And texte.Contains("POTEAU") <> True And texte.Contains("DORMANT") <> True And texte.Contains("SoudureDormant") <> True) Or texte.Contains("SoudureOuvrant") Or texte.Contains("O u v r a n t") Or texte.Contains("O U V R A N T") Or texte.Contains("Ouvrant") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> OUVRANT" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "OUVRANT", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "OUVRANT")
            resultat = GetNumbersForm2(_type_Document, texte, "OUVRANT")

        ElseIf (texte.Contains("FIXE") Or texte.Contains("IMPOSTE") Or texte.Contains("I M P O S T E") Or texte.Contains("F I X E")) And RadioButton7.Checked <> True Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> FIXE" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "FIXE", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "FIXE")
            resultat = GetNumbersForm2(_type_Document, texte, "FIXE")

        ElseIf texte.Contains("FIXE AEV") And RadioButton7.Checked = True Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> FIXE" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "FIXE", i, path)

        ElseIf (texte.Contains("MONTANT") Or texte.Contains("PIVOT") Or texte.Contains("SEUIL") Or texte.Contains("Seuil") Or texte.Contains("TRAVERSE") Or texte.Contains("DORMANT") Or texte.Contains("D O R M A N T") Or texte.Contains("D o r m a n t") Or texte.Contains("SoudureDormant") Or texte.Contains("VENTILATION") Or texte.Contains("CASSETTE") Or texte.Contains("TRAPPE") Or texte.Contains("BASE SUPPORT") Or texte.Contains("PIVOT")) And texte.Contains("Meneau") <> True And (texte.Contains("MN") <> True Or texte.Contains("HOMNIA")) And texte.Contains("POTEAU") <> True Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> SEUIL" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "DORMANT", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "DORMANT")
            resultat = GetNumbersForm2(_type_Document, texte, "DORMANT")

        ElseIf (texte.Contains("MENEAU") Or texte.Contains("Meneau") Or texte.Contains("M e n e a u") Or texte.Contains("M E N E A U") Or texte.Contains("TOLE") Or texte.Contains("MN") Or texte.Contains("Bute")) And (texte.Contains("POTEAU") <> True) Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> MENEAU" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "MENEAU", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "MENEAU")
            resultat = GetNumbersForm2(_type_Document, texte, "MENEAU")


        ElseIf texte.Contains("TUBE") Or texte.Contains("BUTEE") Or texte.Contains("T U B E") Or texte.Contains("T u b e") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> TUBE" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "TUBE", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "TUBE")
            resultat = GetNumbersForm2(_type_Document, texte, "TUBE")

        ElseIf texte.Contains("POTEAU") Or texte.Contains("P O T EA U") Or texte.Contains("P o t e a u") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> POTEAU" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "POTEAU", i, path)
            resultat = GetNumbersForm1(_type_Document, texte, "POTEAU")
            resultat = GetNumbersForm2(_type_Document, texte, "POTEAU")

        ElseIf texte.Contains("CJ") Or texte.Contains("C J") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> CJ" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "CJ", i, path)

        ElseIf texte.Contains("QUINCAILLERIE") Or texte.Contains("Q U I N C A I L L E R I E") Or texte.Contains("Q u i n c a i l l e r i e") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> QUINCAILLERIE" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "QUINCAILLERIE", i, path)

        ElseIf texte.Contains("ALU") Or texte.Contains("A L U") Or texte.Contains("A l u") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> ALU" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "ALU", i, path)

        ElseIf texte.Contains("U 134") Or texte.Contains("U134") Or texte.Contains("134mm") Or texte.Contains("U 1 3 4") Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> U134" + vbCrLf, True)
            resultat = AddDatatoCSVfile(_type_Document, "U134", i, path)



        Else
            resultat = AddDatatoCSVfile(_type_Document, "Page_non_prise", i, path)
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> NON PRISE" + vbCrLf, True)
        End If
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->Fin sous trie, Page : " + i.ToString + ", " + _type_Document + vbCrLf, True)
        Return 1
    End Function

    'Afficher
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Dim p_name As New System.IO.StreamReader(path_name)
            While p_name.Peek() >= 0
                Dim name As String = p_name.ReadLine()
                If name <> "" Then
                    TextBox2.Text = name

                End If
            End While
            p_name.Close()


    End Sub

    Private Function cinquièmetrie(_type_Document As String, path As String, texte As String, i As Integer) As Integer
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->Sous trie ACCESSOIRES, Page : " + i.ToString + vbCrLf, True)

        If texte.Contains("SoudureOuvrant") Or texte.Contains("S o u d u r e O u v r a n t") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure OUVRANT", i, path)

        ElseIf texte.Contains("SoudureFixe") Or texte.Contains("S o u d u r e F i x e") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure FIXE", i, path)

        ElseIf texte.Contains("SoudureDormant") Or texte.Contains("S o u d u r e D o r m a n t") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure DORMANT", i, path)

        ElseIf texte.Contains("SoudureMeneau") Or texte.Contains("S o u d u r e M e n e a u") Then
            If texte.Contains("PAC03512") Or texte.Contains("PAC03513") Or texte.Contains("PAC03514") Or texte.Contains("PAC03515") Or texte.Contains("PAC03516") Or texte.Contains("PAC03517") Or texte.Contains("PAC03518") Or texte.Contains("P A C 0 3 5 1 ") Then

                resultat = AddDatatoCSVfile(_type_Document, "Soudure TUBE", i, path)

            ElseIf texte.Contains("PAC03560") Or texte.Contains("P A C 0 3 5 6 0 ") Then

                resultat = AddDatatoCSVfile(_type_Document, "Soudure BUTEE", i, path)

            Else
                resultat = AddDatatoCSVfile(_type_Document, "Soudure MENEAU", i, path)

            End If


        ElseIf texte.Contains("SoudureTube") Or texte.Contains("S o u d u r e T u b e") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure TUBE", i, path)

        ElseIf texte.Contains("SoudurePoteau") Or texte.Contains("S o u d u r e P o t e a u") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure POTEAU", i, path)

        ElseIf texte.Contains("SoudureCj") Or texte.Contains("S o u d u r e C j") Then
            resultat = AddDatatoCSVfile(_type_Document, "Soudure CJ", i, path)

        ElseIf texte.Contains("PeintureOuvrant") Or texte.Contains("P e i n t u r e O u v r a n t") Then
            resultat = AddDatatoCSVfile(_type_Document, "Peinture OUVRANT", i, path)

        ElseIf texte.Contains("PeintureFixe") Or texte.Contains("P e i n t u r e F i x e") Then
            resultat = AddDatatoCSVfile(_type_Document, "Peinture FIXE", i, path)

        ElseIf texte.Contains("PeintureDormant") Or texte.Contains("P e i n t u r e D o r m a n t") Then
            resultat = AddDatatoCSVfile(_type_Document, "Peinture DORMANT", i, path)

        ElseIf texte.Contains("PeintureMeneau") Or texte.Contains("P e i n t u r e M e n e a u") Then
            resultat = AddDatatoCSVfile(_type_Document, "Peinture MENEAU", i, path)

        ElseIf texte.Contains("PeintureCj") Then
            resultat = AddDatatoCSVfile(_type_Document, "PeinturE CJ", i, path)


        ElseIf texte.Contains("QUINCAILLERIE") Or texte.Contains("Q U I N C A I L L E R I E") Then
            resultat = AddDatatoCSVfile(_type_Document, "QUINCAILLERIE", i, path)

        ElseIf texte.Contains("Expédition") Or texte.Contains("FinitionOuvarant") Or texte.Contains("FinitionDormant") Or texte.Contains("FinitionAlu") Or texte.Contains("E x p é d i t i o n") Or texte.Contains("F i n i t i o n") Then
            condition_Exp = True
            resultat = AddDatatoCSVfile(_type_Document, "Expedition_Finition", i, path)

        ElseIf condition_Exp = True Then
            resultat = AddDatatoCSVfile(_type_Document, "Expedition_Finition", i, path)
        ElseIf texte.Contains("C o n t r o l e L a n c e m e") Or texte.Contains("ControleLanceme") Then

            resultat = AddDatatoCSVfile(_type_Document, "Page_non_prise", i, path)
        Else
            resultat = AddDatatoCSVfile(_type_Document, "NON PRISE", i, path)
        End If
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->Fin sous trie ACCESSOIRES, Page : " + i.ToString + vbCrLf, True)
        Return 1
    End Function

    'Retourne les Ids des plans
    Private Function GetNumbersForm1(_type_Document As String, texte As String, Type As String)
        Dim i As Integer = 0

        Dim theString As String = texte
        Dim numRegEx As New System.Text.RegularExpressions.Regex("[0-9]{8}")


        Dim matches As System.Text.RegularExpressions.MatchCollection = numRegEx.Matches(theString)


        If _type_Document = "Etat_Pieces" Then

            While i < matches.Count

                If Type = "OUVRANT" And _Ids_OUV.Contains(matches(i).Value) <> True Then
                    _Ids_OUV(j__OUV) = matches(i).Value
                    str__OUV = str__OUV + "," + matches(i).Value.ToString
                    j__OUV = j__OUV + 1

                ElseIf Type = "DORMANT" And _Ids_DOR.Contains(matches(i).Value) <> True Then
                    _Ids_DOR(j__DOR) = matches(i).Value
                    str__DOR = str__DOR + "," + matches(i).Value.ToString
                    j__DOR = j__DOR + 1

                ElseIf Type = "POTEAU" And _Ids_POT.Contains(matches(i).Value) <> True Then
                    _Ids_POT(j__POT) = matches(i).Value
                    str__POT = str__POT + "," + matches(i).Value.ToString
                    j__POT = j__POT + 1


                ElseIf Type = "MENEAU" And _Ids_MEN.Contains(matches(i).Value) <> True Then
                    _Ids_MEN(j__MEN) = matches(i).Value
                    str__MEN = str__MEN + "," + matches(i).Value.ToString
                    j__MEN = j__MEN + 1

                ElseIf Type = "FIXE" And _Ids_FIX.Contains(matches(i).Value) <> True Then
                    _Ids_FIX(j__FIX) = matches(i).Value
                    str__FIX = str__FIX + "," + matches(i).Value.ToString
                    j__FIX = j__FIX + 1

                End If


                i = i + 1
            End While

        End If



        Return 1

    End Function

    Private Function GetNumbersForm2(_type_Document As String, texte As String, Type As String)
        Dim i As Integer = 0

        Dim theString As String = texte

        Dim numRegEx As New System.Text.RegularExpressions.Regex("[0-9]{12}")

        Dim matches As System.Text.RegularExpressions.MatchCollection = numRegEx.Matches(theString)

        If _type_Document = "Etat_Pieces" Then

            While i < matches.Count

                If Type = "OUVRANT" And _Ids_OUV.Contains(matches(i).Value) <> True Then
                    _Ids_OUV(j__OUV) = matches(i).Value
                    str__OUV = str__OUV + "," + matches(i).Value.ToString
                    j__OUV = j__OUV + 1

                ElseIf Type = "DORMANT" And _Ids_DOR.Contains(matches(i).Value) <> True Then
                    _Ids_DOR(j__DOR) = matches(i).Value.ToString
                    str__DOR = str__DOR + "," + matches(i).Value.ToString
                    j__DOR = j__DOR + 1

                ElseIf Type = "POTEAU" And _Ids_POT.Contains(matches(i).Value) <> True Then
                    _Ids_POT(j__POT) = matches(i).Value
                    str__POT = str__POT + "," + matches(i).Value.ToString
                    j__POT = j__POT + 1


                ElseIf Type = "MENEAU" And _Ids_MEN.Contains(matches(i).Value) <> True Then
                    _Ids_MEN(j__MEN) = matches(i).Value
                    str__MEN = str__MEN + "," + matches(i).Value.ToString
                    j__MEN = j__MEN + 1

                ElseIf Type = "FIXE" And _Ids_FIX.Contains(matches(i).Value) <> True Then
                    _Ids_FIX(j__FIX) = matches(i).Value
                    str__FIX = str__FIX + "," + matches(i).Value.ToString
                    j__FIX = j__FIX + 1

                End If


                i = i + 1
            End While

        End If

        Return 1

    End Function '

    Private Function AddHeader(service As String, page_service As Integer, reader4 As PdfReader) As Integer


        '''''''''''''''''''''''''''''''''''''''En tête'''''''''''''''''''''''''''''''''''''''''''''''''

        'Dim reader4 As New PdfReader(Path) ' Fichier final crée

        'Dim size As Rectangle = reader4.GetPageSizeWithRotation(1)
        Dim document2 As New Document

        ' Create the writer
        Dim fs As New FileStream("C:\Traitement DiffusionDocsFab\" + service + ".PDF", FileMode.Create, FileAccess.Write)
        Dim writer2 As PdfWriter = PdfWriter.GetInstance(document2, fs)
        document2.Open()
        Dim cb As PdfContentByte = writer2.DirectContent

        ' Set the font, color and size properties for writing text to the PDF
        Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
        cb.SetColorFill(BaseColor.BLACK)
        cb.SetFontAndSize(bf, 14)

        ' Write text in the PDF
        cb.BeginText()
        Dim text As String = "      " + service + "         "
        Dim text2 As String = TextBox2.Text

        ' Set the alignment and coordinates here
        '300, 810, 0
        cb.ShowTextAligned(1, text, 560, 830, 0)
        cb.ShowTextAligned(1, text2, 70, 830, 0)
        cb.EndText()

        ' Put the text on a new page in the PDF

        Dim page As PdfImportedPage = writer2.GetImportedPage(reader4, page_service)

        cb.AddTemplate(page, 0, 0)

        ' Close the objects

        document2.Close()
        fs.Close()
        writer2.Close()
        'reader4.Close()

        Return 1

    End Function

    Private Function GetDataFromCsv(_Document As String, _Composant As String, _service As String)
        Dim lignes() As String = System.IO.File.ReadAllLines(_strCustomerCSVPath)
        Dim i As Integer
        Dim condition_EP As Boolean = False



        Select Case _Document

            'Cas Etat des pièces 

            Case "Etat_Pieces"
                If _comp_ligne(0) <> 0 Then
                    For i = _comp_ligne(0) - 1 To _comp_ligne(1) - 1

                        If Split(lignes(i), ",")(1) = _Composant Then

                            If condition_EP = False And _Entete_service(0) <> 0 Then

                                condition_EP = True
                                Dim reader As New PdfReader(Split(lignes(i), ",")(3))

                                resultat = AddHeader2(_Composant, Split(lignes(i), ",")(2), reader)
                                Dim path_EP = "C:\Traitement DiffusionDocsFab\" + _Composant + ".PDF"
                                Dim reader_addpages As New PdfReader(path_EP)
                                Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                                copier.AddPage(importedPage)

                                _cmpt_total = _cmpt_total + 1
                                _Pages_Etat(_indice_Pages_Etat) = _cmpt_total
                                _indice_Pages_Etat = _indice_Pages_Etat + 1
                                reader_addpages.Close()

                            Else
                                If _Entete_service(0) = 0 Then

                                    resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))

                                    _Pages_Etat(_indice_Pages_Etat) = _cmpt_total
                                    _indice_Pages_Etat = _indice_Pages_Etat + 1

                                Else
                                    resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                                End If




                            End If

                        End If

                    Next
                End If
            Case "Bordereau_Fabrication"
                If _comp_ligne(2) <> 0 Then
                    For i = _comp_ligne(2) - 1 To _comp_ligne(3) - 1

                        If Split(lignes(i), ",")(1) = _Composant Then

                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                            If (_service = "S10" Or _service = "S30") And (RadioButton2.Checked = True Or RadioButton3.Checked = True) Then

                                _Pages_Etat(_indice_Pages_Etat) = _cmpt_total
                                _indice_Pages_Etat = _indice_Pages_Etat + 1

                            End If

                        End If

                    Next
                End If

            Case "Accessoires_Exp"
                If _comp_ligne(4) <> 0 Then
                    If _comp_ligne(4) < _comp_ligne(5) Then
                        For i = _comp_ligne(4) - 1 To _comp_ligne(5) - 1
                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        Next
                    Else
                        i = _comp_ligne(4) - 1
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    End If
                End If

            Case "Accessoires"
                If _comp_ligne(6) <> 0 Then
                    If _comp_ligne(6) < _comp_ligne(7) Then

                        For i = _comp_ligne(6) - 1 To _comp_ligne(7) - 1
                            If Split(lignes(i), ",")(1).StartsWith(_Composant) Then

                                resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                            End If

                        Next
                    Else
                        i = _comp_ligne(6) - 1
                        If Split(lignes(i), ",")(1).StartsWith(_Composant) Then

                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        End If
                    End If
                End If


            Case "BT ALUMINIUM"
                If _comp_ligne(8) <> 0 Then
                    If _comp_ligne(8) < _comp_ligne(9) Then

                        For i = _comp_ligne(8) - 1 To _comp_ligne(9) - 1
                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                            If RadioButton2.Checked = True Or RadioButton3.Checked = True Then
                                _Pages_Etat(_indice_Pages_Etat) = _cmpt_total
                                _indice_Pages_Etat = _indice_Pages_Etat + 1

                            End If
                        Next
                    Else
                        i = _comp_ligne(8) - 1
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        If RadioButton2.Checked = True Or RadioButton3.Checked = True Then
                            _Pages_Etat(_indice_Pages_Etat) = _cmpt_total
                            _indice_Pages_Etat = _indice_Pages_Etat + 1

                        End If
                    End If
                End If


            Case "Liste_Encadrement"
                If _comp_ligne(10) <> 0 Then
                    If _comp_ligne(10) < _comp_ligne(11) Then

                        For i = _comp_ligne(10) - 1 To _comp_ligne(11) - 1

                            If Split(lignes(i), ",")(0).StartsWith("Liste") Then

                                resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                            End If

                        Next
                    Else
                        i = _comp_ligne(10) - 1
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    End If
                End If


            Case "Vide"
                If _comp_ligne(12) <> 0 Then
                    If _comp_ligne(12) < _comp_ligne(13) Then

                        For i = _comp_ligne(12) - 1 To _comp_ligne(13) - 1
                            If Split(lignes(i), ",")(0) = "Vide" Then

                                resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))

                            End If
                        Next
                    Else
                        i = _comp_ligne(12) - 1
                        If Split(lignes(i), ",")(0) = "Vide" Then

                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))

                        End If
                    End If
                End If

            Case "SOUDAGE"

                For i = _cmpt_SOUDAGE(0) - 1 To _cmpt_SOUDAGE(1) - 1

                    If Split(lignes(i), ",")(0) = "SOUDAGE" And Split(lignes(i + 1), ",")(0) = "Vide" Then

                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        resultat = AddPages(Split(lignes(i + 1), ",")(3), _service, Split(lignes(i + 1), ",")(2))

                    ElseIf Split(lignes(i), ",")(0) = "SOUDAGE" Then
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    End If

                Next

            Case "MAGASIN"

                For i = _cmpt_MAGASIN(0) - 1 To _cmpt_MAGASIN(1) - 1

                    If Split(lignes(i), ",")(0) = "MAGASIN" And Split(lignes(i + 1), ",")(0) = "Vide" Then

                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        resultat = AddPages(Split(lignes(i + 1), ",")(3), _service, Split(lignes(i + 1), ",")(2))

                    ElseIf Split(lignes(i), ",")(0) = "MAGASIN" Then
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    End If

                Next

            Case "PEINTURE"

                For i = _cmpt_PEINTURE(0) - 1 To _cmpt_PEINTURE(1) - 1

                    If Split(lignes(i), ",")(0) = "PEINTURE" Then

                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))

                    End If

                Next

            Case "ALU"

                For i = _cmpt_ALU(0) - 1 To _cmpt_ALU(1) - 1

                    If Split(lignes(i), ",")(0) = "ALU" Then

                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    End If

                Next

            Case "Plan_S60"

                If _cmpt_PLANS_S60(0) <> 0 Then
                    For i = _cmpt_PLANS_S60(0) - 1 To _cmpt_PLANS_S60(1) - 1

                        If Split(lignes(i), ",")(0) = "Plan_S60" Then
                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        End If
                    Next
                End If

                'Plan Tube
            Case "Plan_TUBE"

                If _cmpt_PLANS_TUBE(0) <> 0 Then
                    For i = _cmpt_PLANS_TUBE(0) - 1 To _cmpt_PLANS_TUBE(1) - 1
                        resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                    Next
                End If


                'Plan S00
            Case "Plan_S00"

                If _cmpt_PLANS_S00(0) <> 0 Then
                    For i = _cmpt_PLANS_S00(0) - 1 To _cmpt_PLANS_S00(1) - 1

                        If Split(lignes(i), ",")(1) = _Composant Then

                            resultat = AddPages(Split(lignes(i), ",")(3), _service, Split(lignes(i), ",")(2))
                        End If


                    Next
                End If


        End Select
        Return 1
    End Function
    Private Function AddPages(_Path_Reader As String, _service As String, _page As Integer) As Integer
        Dim reader As New PdfReader(_Path_Reader)
        doc.Open()

        _cmpt_total = _cmpt_total + 1

        Select Case _service

            Case "S00"
                If _Entete_service(0) = 0 Then
                    _Entete_service(0) = _page
                    resultat = AddHeader("00", _Entete_service(0), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "00" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)
                    reader_addpages.Close()
                Else

                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If

            Case "S10"
                If _Entete_service(1) = 0 Then
                    _Entete_service(1) = _page
                    resultat = AddHeader("10", _Entete_service(1), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "10" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)



                End If

            Case "S20"
                If _Entete_service(2) = 0 Then
                    _Entete_service(2) = _page
                    resultat = AddHeader("20", _Entete_service(2), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "20" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If
            Case "S30"
                If _Entete_service(3) = 0 Then
                    _Entete_service(3) = _page
                    resultat = AddHeader("30", _Entete_service(3), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "30" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If
            Case "S40"
                If _Entete_service(4) = 0 Then
                    _Entete_service(4) = _page
                    resultat = AddHeader("40", _Entete_service(4), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "40" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If
            Case "S50"
                If _Entete_service(5) = 0 Then
                    _Entete_service(5) = _page
                    resultat = AddHeader("50", _Entete_service(5), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "50" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If
            Case "S60"
                If _Entete_service(6) = 0 Then
                    _Entete_service(6) = _page
                    resultat = AddHeader("60", _Entete_service(6), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "60" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)



                End If
            Case "S70"
                If _Entete_service(7) = 0 Then
                    _Entete_service(7) = _page
                    resultat = AddHeader("70", _Entete_service(7), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "70" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)



                End If

            Case "S90"
                If _Entete_service(8) = 0 Then
                    _Entete_service(8) = _page
                    resultat = AddHeader("90", _Entete_service(8), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "90" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If

            Case "S100"
                If _Entete_service(9) = 0 Then
                    _Entete_service(9) = _page
                    resultat = AddHeader("100", _Entete_service(9), reader)
                    Dim path_service = "C:\Traitement DiffusionDocsFab\" + "100" + ".PDF"
                    Dim reader_addpages As New PdfReader(path_service)
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader_addpages, 1)
                    copier.AddPage(importedPage)

                    reader_addpages.Close()
                Else
                    Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                    copier.AddPage(importedPage)


                End If

            Case Else

                Dim importedPage As PdfImportedPage = copier.GetImportedPage(reader, _page)
                copier.AddPage(importedPage)
        End Select




        'copier.Close()
        'doc.Close()
        reader.Close()
        Return 1
    End Function

    Private Function AddPlans(_composant_id As String, _service_plans As String, _name As String)

        Dim parts(5) As String

        If repitition.Contains(_composant_id) <> True Then
            repitition = repitition + " , " + _composant_id
            Select Case Len(_composant_id.ToString)
            '010501
                Case 6

                    parts(0) = _composant_id.Substring(0, 2)
                    parts(1) = _composant_id.Substring(2, 2)
                    parts(2) = _composant_id.Substring(4, 2)
                    'MsgBox(parts(0) & "," & parts(1) & "," & parts(2))
                    pp_plan = 0
                    While Names_Files_plans(pp_plan) <> ""
                        If Names_Files_plans_prises.Contains(Names_Files_plans(pp_plan)) <> True And Names_Files_plans(pp_plan).Contains("PAC02601") <> True Then


                            If Split(Names_Files_plans(pp_plan), source)(1).Contains("-" + parts(0) + "_" + parts(1) + "_" + parts(2)) = True And Names_Files_plans_prises.Contains(Names_Files_plans(pp_plan)) <> True And Split(Names_Files_plans(pp_plan), source)(1).Contains("-" + parts(0) + "_" + parts(1) + "_" + parts(2) + "-") <> True Then

                                'MsgBox(parts(0) & "," & parts(1) & "," & parts(2))
                                _nombre_plans = _nombre_plans + 1
                                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + Names_Files_plans(pp_plan) + " a été pris" + vbCrLf, True)
                                AddPages(Names_Files_plans(pp_plan), _service_plans, 1)
                                Names_Files_plans_prises = Names_Files_plans_prises + " , " + Names_Files_plans(pp_plan)

                            End If
                        End If
                        pp_plan = pp_plan + 1
                    End While
            End Select
        End If

        Return 1

    End Function

    Private Function Fusion(_Path_Reader As String) As Integer


        Dim _page As Integer
        Dim reader As New PdfReader(_Path_Reader)
        Dim numberOfPages As Integer = reader.NumberOfPages
        doc_non_trie.Open()

        For _page = 1 To numberOfPages

            Dim importedPage As PdfImportedPage = copier_non_trie.GetImportedPage(reader, _page)
            copier_non_trie.AddPage(importedPage)
        Next


        reader.Close()
        Return 1
    End Function

    Private Sub CloseMe()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf CloseMe))
            Exit Sub
        End If
        Me.Close()
    End Sub

    Private Function Correction_plans(_name As String, _service_plans As String) As Integer
        pp_plan = 0
        While Names_Files_plans(pp_plan) <> ""
            If Names_Files_plans_prises.Contains(Names_Files_plans(pp_plan)) <> True And Names_Files_plans(pp_plan).Contains("PAC02601") <> True Then

                Dim pr As New PdfReader(Names_Files_plans(pp_plan))
                If Scan(Names_Files_plans(pp_plan), 1).contains(_name) And Scan(Names_Files_plans(pp_plan), 1).contains("ALU") <> True Then

                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "Le fichier : " + Names_Files_plans(pp_plan) + " a été pris -> Correction" + vbCrLf, True)
                    AddPages(Names_Files_plans(pp_plan), _service_plans, 1)
                    Names_Files_plans_prises = Names_Files_plans_prises + " , " + Names_Files_plans(pp_plan)

                End If
            End If
            pp_plan = pp_plan + 1
        End While
        Return 1
    End Function

    Private Function AddHeader2(service As String, page_service As Integer, reader4 As PdfReader) As Integer


        '''''''''''''''''''''''''''''''''''''''En tête'''''''''''''''''''''''''''''''''''''''''''''''''

        'Dim size As Rectangle = reader4.GetPageSizeWithRotation(1)
        Dim document2 As New Document

        ' Create the writer
        Dim fs As New FileStream("C:\Traitement DiffusionDocsFab\" + service + ".PDF", FileMode.Create, FileAccess.Write)
        Dim writer2 As PdfWriter = PdfWriter.GetInstance(document2, fs)
        document2.Open()
        Dim cb As PdfContentByte = writer2.DirectContent

        ' Set the font, color and size properties for writing text to the PDF
        Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
        cb.SetColorFill(BaseColor.BLACK)
        cb.SetFontAndSize(bf, 14)

        ' Write text in the PDF
        cb.BeginText()

        Dim text2 As String = TextBox2.Text


        ' Set the alignment and coordinates here
        '300, 810, 0

        cb.ShowTextAligned(1, text2, 70, 830, 0)
        cb.EndText()

        ' Put the text on a new page in the PDF

        Dim page As PdfImportedPage = writer2.GetImportedPage(reader4, page_service)
        cb.AddTemplate(page, 0, 0)

        ' Close the objects

        document2.Close()
        fs.Close()
        writer2.Close()
        'reader4.Close()

        Return 1

    End Function

    Private Function Inistialisation() As Integer
        source = ""
        resultat = 0
        _strCustomerCSVPath = ""
        Names_Files_plans_prises = ""

        j__OUV = 0
        j__DOR = 0
        j__POT = 0
        j__MEN = 0
        j__FIX = 0

        str__OUV = ","
        str__DOR = ","
        str__POT = ","
        str__MEN = ","
        str__FIX = ","

        _nombre_plans = 0
        _comp = 0
        p_plan = 0
        pp_plan = 0
        repitition = ""

        condition_Exp = False
        year = ""


        For i As Integer = 0 To 300
            Names_Files(i) = ""
            Names_Files_plans(i) = ""
            _Ids_OUV(i) = 0
            _Ids_DOR(i) = 0
            _Ids_POT(i) = 0
            _Ids_MEN(i) = 0
            _Ids_FIX(i) = 0
        Next


        For i As Integer = 0 To 15
            Names_Files(i) = ""
            Names_Files_plans(i) = ""
            _comp_ligne(i) = 0
            _Entete_service(i) = 0
            _Condition_Cib(i) = False

        Next



        Return 1
    End Function

    Private Function WaterMark(_path As String)

        Dim reader As New PdfReader(_path)
        Dim fs As New FileStream("C:\Traitement DiffusionDocsFab\" + TextBox1.Text + "_Pôles.PDF", FileMode.Create, FileAccess.Write)
        Dim pdfStamper As New PdfStamper(reader, fs)
        Dim pathpic As String
        Dim i As Integer = 0

        If RadioButton2.Checked = True Or RadioButton3.Checked = True Or RadioButton8.Checked = True Then
            'PEINTURE LISSE
            pathpic = "\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\PEINTURE_LISSE.png"

        ElseIf RadioButton4.Checked = True Then

            'INOX SB
            pathpic = "\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\INOX_SCOTCH_BRIT.png"

        ElseIf RadioButton5.Checked = True Then
            'INOX PP
            pathpic = "\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\INOX_PEAU_DE_PORC.png"

        ElseIf RadioButton6.Checked = True Then
            'INOX DM
            pathpic = "\\srvdc\Bureau_Etudes\instal_BE\Cibox2022 DocsFab\INOX_DAMIER.png"

        End If


        Dim image = iTextSharp.text.Image.GetInstance(pathpic)

        While _Pages_Etat(i) <> 0
            Dim content = pdfStamper.GetOverContent(_Pages_Etat(i))

            image.SetAbsolutePosition(200, 635)
            content.AddImage(image)



            image.SetAbsolutePosition(200, 370)
            content.AddImage(image)

            i = i + 1

        End While

        pdfStamper.Close()


    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''CibSlide''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Code ajouté pour le trie de la CibSlide !

    Private Function _CibSlide_Premiertrie() As Integer
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim _Export_condition As Boolean = False

        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "---> 1/Premier trie" + vbCrLf, True)
        For Each files As String In System.IO.Directory.GetFiles(source)

            ' Les Plans

            If files.Contains(".pdf") And (Split(files, "\")(6).StartsWith("D") <> True And Split(files, "\")(6).StartsWith("d") <> True) Then
                resultat = Fusion(files) 'Fusion des fichiers non triés
                Names_Files_plans(j) = files.ToString

                j = j + 1
            ElseIf files <> "" And files.Contains(".pdf") Then 'Doc 1 et 2
                Names_Files(i) = files.ToString
                i = i + 1
            End If

        Next
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "---> 1/Fin Premier trie" + vbCrLf, True)
        Return 1
    End Function


    Private Function _Cibslide_Deuxiemetrie() As Integer

        Dim i = 0
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------> 2/Deuxieme trie" + vbCrLf, True)
        While Names_Files(i) <> ""

            resultat = Fusion(Names_Files(i)) ' Fusion des fichiers non triés
            resultat = _Cibslide_TroisièmeTrie(Names_Files(i))
            i = i + 1
        End While
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------> 2/Fin Deuxieme trie" + vbCrLf, True)
        Return 1
    End Function


    'Trie du fichier Docs
    Private Function _Cibslide_TroisièmeTrie(path_file As String) As Integer
        Dim condition As Boolean = False
        Dim pr As New PdfReader(path_file)
        Dim numberOfPages As Integer = pr.NumberOfPages
        Dim str As String = ""

        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------> 3/Troisième trie : " + path_file + vbCrLf, True)
        For i = 1 To numberOfPages

            str = Scan(path_file, i)
            'MsgBox(str)

            If str.Contains("ENCADREMENTS") Then
                resultat = AddDatatoCSVfile("Liste_Encadrement", "Vide", i, path_file)
            ElseIf str.Contains("BT ALUMINIUM") Then
                resultat = AddDatatoCSVfile("BT ALUMINIUM", "Vide", i, path_file)

            ElseIf str.Contains("Etat") Or str.Contains("Pièces") Then
                resultat = Quatrièmetrie("Etat_Pieces", path_file, str, i)
                'resultat = AddDatatoCSVfile("Etat_Pieces", "Vide", i, path_file)

            ElseIf str.Contains("Accessoires") Or str.Contains("A c c e s s o i r e s") Or str.Contains("P r é p a r a t i o n") Then
                'resultat = AddDatatoCSVfile("Accessoires", "Vide", i, path_file)
                resultat = cinquièmetrie("Accessoires", path_file, str, i)

            ElseIf str.Contains("Bodereau") Or str.Contains("Fabrication") Or str.Contains("F a b r i c a t i o n") Then

                'resultat = AddDatatoCSVfile("Bordereau_Fabrication", "Vide", i, path_file)
                resultat = Quatrièmetrie("Bordereau_Fabrication", path_file, str, i)
            Else
                resultat = AddDatatoCSVfile("Vide", "Pages_non_prise", i, path_file)
            End If

        Next
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------> 3/Fin Troisième trie : " + path_file + vbCrLf, True)
        Return 1

    End Function


    Private Function _Cibslide_Cinquièmetrie() As Integer

        Dim i = 0

        XLS1 = New Excel.Application()

        If XLS1 Is Nothing Then

            MsgBox("Excel n'est pas installé!!" + vbCrLf + " Reexécuter la fusion une fois le problème réglé")

        End If
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------> 4/Plan trie" + vbCrLf, True)

        Dim misValue As Object = System.Reflection.Missing.Value
        Classeur1 = XLS1.Workbooks.Add(misValue)

        Classeur1.SaveAs(destExcelFileRoot_FeuilleComposant, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Classeur1.Close(True, misValue, misValue)

        While Names_Files_plans(i) <> ""


            resultat = _Cibslide_PlansTrie(Names_Files_plans(i))
            i = i + 1
        End While

        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------> 4/Fin Plan trie" + vbCrLf, True)
        Return 1
    End Function

    Private Function _Cibslide_PlansTrie(path_file As String) As Integer
        Dim pr As New PdfReader(path_file)
        Dim numberOfPages As Integer = pr.NumberOfPages
        Dim str As String = ""
        Dim condition_lot As Boolean = True
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> 5/Plan : " + path_file + vbCrLf, True)


        For i = 1 To numberOfPages

            str = Scan(path_file, i)

            If (str.Contains("Aluminium") And str.Contains("LISTE") <> True) Or (str.Contains("CAV") And str.Contains("TYPE")) Then
                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->Aluminium, Page : " + i.ToString + vbCrLf, True)
                If str.Contains("RAIL MOTEUR EQUIPE") Then
                    If _cmpt_PLANS_S60(0) = 0 Then
                        _cmpt_PLANS_S60(0) = _comp + 1
                    End If
                    _cmpt_PLANS_S60(1) = _comp + 1

                    resultat = AddDatatoCSVfile("Plan_S60", "MONTAGE", i, path_file)
                End If


                If _cmpt_ALU(0) = 0 Then
                    _cmpt_ALU(0) = _comp + 1
                End If
                _cmpt_ALU(1) = _comp + 1
                resultat = AddDatatoCSVfile("ALU", "ALU", i, path_file)


            ElseIf str.Contains("RAIL MOTEUR EQUIPE") Then
                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->RAIL MOTEUR EQUIPE, Page : " + i.ToString + vbCrLf, True)
                If _cmpt_PLANS_S60(0) = 0 Then
                    _cmpt_PLANS_S60(0) = _comp + 1
                    'MsgBox(_cmpt_PLANS_S60(0))
                End If
                _cmpt_PLANS_S60(1) = _comp + 1
                'MsgBox(_cmpt_PLANS_S60(1))
                resultat = AddDatatoCSVfile("Plan_S60", "MONTAGE", i, path_file)

            ElseIf str.Contains("TUBE") Then

                My.Computer.FileSystem.WriteAllText(_Fichier_Log, "----------------------------------------------------------->TUBE, Page : " + i.ToString + vbCrLf, True)
                If _cmpt_PLANS_TUBE(0) = 0 Then
                    _cmpt_PLANS_TUBE(0) = _comp + 1

                End If
                _cmpt_PLANS_TUBE(1) = _comp + 1

                resultat = AddDatatoCSVfile("Plan_TUBE", "ATELIER", i, path_file)

            Else


                resultat = _Cibslide_ExportTrie(pathExportVault, i, path_file)

            End If
        Next
        My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> 5/Fin Plan : " + path_file + vbCrLf, True)

        Return 1
    End Function

    Private Function _Cibslide_ExportTrie(path_export As String, i As Integer, path_file As String) As Integer
        Dim ir As Integer
        Dim found As Boolean = False
        Dim lignes As String() = System.IO.File.ReadAllLines(path_export)
        Dim numplan As String


        numplan = Split(Split(path_file, "\")(6), ".pdf")(0)

        For ir = 0 To lignes.Count - 1

            If Split(lignes(ir), ",")(2) = numplan Then
                found = True
                If _cmpt_PLANS_S00(0) = 0 Then
                    _cmpt_PLANS_S00(0) = _comp + 1

                End If
                _cmpt_PLANS_S00(1) = _comp + 1

                resultat = AddDatatoCSVfile("Plan_S00", Split(lignes(ir), ",")(7), i, path_file)

                Classeur1 = XLS1.Workbooks.Open(destExcelFileRoot_FeuilleComposant, 0, False, 5, "", "", False, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", True, False, 0, True, False, False)

                If listecomp.Contains(Split(lignes(ir), ",")(7)) <> True Then
                    My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> Export, Page : " + i.ToString + " Composant : " + Split(lignes(ir), ",")(7) + vbCrLf, True)
                    Classeur1.Worksheets.Add().Name = Split(lignes(ir), ",")(7)
                    Classeur1.Worksheets(Split(lignes(ir), ",")(7)).Cells(1, 1).value = "Composant : " + vbCrLf + Split(lignes(ir), ",")(7)
                    Classeur1.Worksheets(Split(lignes(ir), ",")(7)).Cells(1, 1).Font.Size = 72
                    Classeur1.Worksheets(Split(lignes(ir), ",")(7)).Cells(1, 1).Font.Bold = True
                    Classeur1.Worksheets(Split(lignes(ir), ",")(7)).columns("A").ColumnWidth = 100

                    With Classeur1.Worksheets(Split(lignes(ir), ",")(7)).Range("A:B")

                        .HorizontalAlignment = Excel.Constants.xlCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                    End With


                    Classeur1.Worksheets(Split(lignes(ir), ",")(7)).ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    "C:\Traitement DiffusionDocsFab\Feuille " + Split(lignes(ir), ",")(7) + ".pdf",
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        True,
                        True,
                        1,
                        10,
                        False)

                    listecomp(inlistecomp) = Split(lignes(ir), ",")(7)
                    inlistecomp = inlistecomp + 1

                End If
                Classeur1.Save()
                Classeur1.Close()
                Exit For
            End If

        Next


        If found = False Then
            My.Computer.FileSystem.WriteAllText(_Fichier_Log, "-----------------------------------------------------------> Export Non renseigné, Page : " + i.ToString + vbCrLf, True)
            resultat = AddDatatoCSVfile("Plan_S00", "Non renseigne", i, path_file)
        End If

        Return 1
    End Function


End Class
