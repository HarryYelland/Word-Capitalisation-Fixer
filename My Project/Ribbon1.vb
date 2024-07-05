Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

    End Sub

    Private ReadOnly wordsToCapitalize As List(Of String) = New List(Of String) From {
"Oliver", "George", "Harry", "Jack", "Jacob", "Charlie", "Thomas", "Henry", "Oscar", "William", "James", "Leo", "Joshua", "Freddie", "Alfie", "Archie", "Ethan", "Isaac", "Alexander", "Joseph", "Samuel", "Sebastian", "Edward", "Arthur", "Logan", "Harrison", "Daniel", "Theo", "Matthew", "Lucas", "Lewis", "Finn", "Hugo", "Adam", "Dylan", "Zachary", "Rory", "Reuben", "Benjamin", "Jake", "Max", "Elijah", "Mason", "Ryan", "Nathan", "Toby", "Frankie", "Gabriel", "Theodore", "David", "Bobby", "Harvey", "Caleb", "Elliot", "Albie", "Jude", "Luke", "Michael", "Elliott", "Ronnie", "Stanley", "Louis", "Finley", "Jasper", "Liam", "Jamie", "Jenson", "Ralph", "Patrick", "Ezra", "Myles", "Hudson", "Ruben", "Milo", "Arlo", "Grayson", "Hunter", "Roman", "Rowan", "Reggie", "Alex", "Blake", "Charles", "Jackson", "Austin", "Carter", "Jesse", "Muhammad", "Aidan", "Felix", "Albert", "Muhammad", "Muhammad",
"London", "Birmingham", "Manchester", "Liverpool", "Leeds", "Sheffield", "Edinburgh", "Glasgow", "Bristol", "Cardiff", "Cambridge", "Oxford", "York", "Bath", "Newcastle", "Brighton", "Southampton", "Belfast", "Dublin", "Microsoft", "Google", "Apple", "Amazon", "Facebook", "IBM", "Intel", "Samsung", "Sony", "Cisco", "Toyota", "Volkswagen", "Coca-Cola", "Pepsi", "McDonald's", "Nike", "Adidas", "Honda", "BMW", "Mercedes"}


    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim selection As Word.Selection = Globals.ThisAddIn.Application.Selection

        If selection IsNot Nothing AndAlso selection.Range IsNot Nothing Then
            Dim selectedText As String = selection.Text
            Dim processedText As String = ConvertTextWithCapitalizedWords(selectedText)

            selection.Text = processedText
        End If
    End Sub

    Private Function ConvertTextWithCapitalizedWords(ByVal text As String) As String
        Dim delimiters As Char() = {" "c, "."c, ","c, ";"c, ":"c, "!"c, "?"c, vbCrLf}
        Dim sentenceEndings As Char() = {"."c, "!"c, "?"c}
        Dim wordsAndDelimiters As New List(Of String)
        Dim currentWord As New StringBuilder()

        For Each ch As Char In text
            If delimiters.Contains(ch) Then
                If currentWord.Length > 0 Then
                    wordsAndDelimiters.Add(currentWord.ToString())
                    currentWord.Clear()
                End If
                wordsAndDelimiters.Add(ch.ToString())
            Else
                currentWord.Append(ch)
            End If
        Next

        If currentWord.Length > 0 Then
            wordsAndDelimiters.Add(currentWord.ToString())
        End If

        Dim capitalizeNextWord As Boolean = True

        For i As Integer = 0 To wordsAndDelimiters.Count - 1
            Dim word As String = wordsAndDelimiters(i)
            If Not delimiters.Contains(word(0)) Then
                If capitalizeNextWord Then
                    wordsAndDelimiters(i) = Char.ToUpper(word(0)) & word.Substring(1).ToLower()
                    capitalizeNextWord = False
                Else
                    If wordsToCapitalize.Any(Function(w) String.Equals(w, word, StringComparison.OrdinalIgnoreCase)) Then
                        wordsAndDelimiters(i) = wordsToCapitalize.Find(Function(w) String.Equals(w, word, StringComparison.OrdinalIgnoreCase))
                    Else
                        wordsAndDelimiters(i) = word.ToLower()
                    End If
                End If
            End If

            If sentenceEndings.Contains(word(0)) Then
                capitalizeNextWord = True
            End If
        Next

        Return String.Join("", wordsAndDelimiters)
    End Function

End Class
