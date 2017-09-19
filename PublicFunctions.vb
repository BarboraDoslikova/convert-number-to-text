Imports Microsoft.VisualBasic
Imports SqlFunctions ' to load data to convertNumberToText() function
Imports System.Data ' for work with DataSet class in convertNumberToText() function
Imports System.Collections.Generic ' for List(Of T)
Imports System

' Class: noun; TextFunctions
' Variable: noun; bookingHistory
' Method: verb; convertNumberToText()
' Constant: MIN_REGISTRATION_AGE

Public Class TextFunctions ' because defaults to private, only public functions regarding returning text



    Shared Function convertNumberToText(ByVal i As Double) As String ' shared so no need to instantiate, just call the function; public by default
        ' Funtion summary:
        ' accepts a number of 1-9 integers and any number of decimal places, e.g. 1000,12999999
        ' returns the number as (Czech) text, e.g. jeden tisíc korun českých a třináct haléřů

        Dim inputNumberDb As Double = i ' accept number from the input box, e.g. 1000

        Dim inputNumberIntegerPortionDb As Double = Fix(inputNumberDb) ' from 1234.9 -> 1234

        Dim inputNumberRoundedRemainderDb As Double = Math.Round(inputNumberDb Mod 1, 2) ' from 1234.129 -> 0.13	

        Dim returnWordStr As String ' return word, e.g. jeden tisíc korun českých

        ' CHECKS the inputNumber
        If inputNumberDb = 0 Then
            returnWordStr = "nula korun českých"
            Return returnWordStr
        ElseIf inputNumberDb > 999999999 Then
            returnWordStr = "příliš velké číslo"
            Return returnWordStr
        Else
            ' MAIN FUNCTION for the integer portion of the number (the remainder portion of the number is dealt with separately, as will not be used often)
            ' convert number to digits
            Dim inputDigitsIntAr As Integer() = Array.ConvertAll(Of Char, Integer)(inputNumberIntegerPortionDb.ToString.ToCharArray, Function(c As Char) Integer.Parse(c.ToString)) '/ {1,2,3,4}
            Array.Reverse(inputDigitsIntAr)

            ' initialize an empty array of the length of the maximum number e.g. 999,999,999
            Dim myDigitsIntAr As Integer() = {0, 0, 0, 0, 0, 0, 0, 0, 0}

            ' assign number digits to an empty array
            For j = 0 To inputDigitsIntAr.Length - 1
                myDigitsIntAr.SetValue(inputDigitsIntAr(j), j)
            Next
            Array.Reverse(myDigitsIntAr) '/ {0,0,1,2,3,4}

            ' initialize an empty array of the length of the maximum number e.g. 999,999
            Dim myWordsStrAr As String() = {"", "", "", "", "", "", "", "", ""}

            ' mapping
            Dim dataset As New DataSet()
            dataset = GetFullSqlData()
            ' this is how to work with SQL data:
            ' e.g. Label1.Text = dataset.Tables(0).Rows(0).Item(2) 'Rows() = 0-based index/row, Item = 0-based column where 0 is the index column ans 1 is the first data column

            Dim hundredsStrAr As String() = {"jedno sto", "dvě sta", "tři sta",
                        "čtyři sta", "pět set", "šest set", "sedm set", "osm set", "devět set"}
            Dim tensStrAr As String() = {"deset", "dvacet", "třicet",
                        "čtyřicet", "padesát", "šedesát", "sedmdesát", "osmdesát", "devadesát"}
            Dim teensStrAr As String() = {"jedenáct", "dvanáct", "třináct",
                        "čtrnáct", "patnáct", "šestnáct", "sedmnáct", "osmnáct", "devatenáct"}
            Dim singlesStrAr As String() = {"jedna", "dva", "tři",
                        "čtyři", "pět", "šest", "sedm", "osm", "devět"}

            ' MATCH digits to mapping / e.g. {" "," ","jeden","dve sta","tricet","ctyri"}		
            For i = 0 To myDigitsIntAr.Length - 1 ' will cycle throught all digits 1st[0] to 6th[5]

                If myDigitsIntAr(i) = 0 Then ' if digit in a position is zero, do nothing
                    'do nothing
                Else ' if digit in a position is anything other than zero, 
                    Dim insertPosition As Integer = i ' the last possible position is 8th, indexed from 0
                    Dim insertWord As String ' the word to be added to the emptry myWordsStrAr array

                    If i = 0 Or i = 3 Or i = 6 Then
                        'insertWord = hundredsStrAr(myDigitsIntAr(i) - 1) ' find word from array's position x-1 as indexed from 0	' do NOT need if loading from SQL!
                        insertWord = dataset.Tables(0).Rows(myDigitsIntAr(i) - 1).Item(1) ' DO THE SAME FOR THE REST OF THE CODE

                    ElseIf i = 1 Or i = 4 Or i = 7 Then
                        If myDigitsIntAr(i) = 1 Then
                            'do nothing
                        Else
                            insertWord = tensStrAr(myDigitsIntAr(i) - 1)
                        End If
                    ElseIf i = 2 Or i = 5 Or i = 8 Then
                        If myDigitsIntAr(i - 1) = 1 Then
                            insertWord = teensStrAr(myDigitsIntAr(i) - 1)
                        Else
                            insertWord = singlesStrAr(myDigitsIntAr(i) - 1)
                        End If
                    End If

                    myWordsStrAr.SetValue(insertWord, insertPosition) ' assign word into the position
                    insertWord = "" ' reset insert word or it inserts hundreds if current teens doesn't return any value
                End If
            Next

            If inputNumberIntegerPortionDb = 2 Then ' handles the one exception for when the inputNumberDb = 2
                myWordsStrAr = {"", "", "", "", "", "", "", "", "dvě"}
            End If


            ' ADD tisic or tisice		
            Dim thousandsToAddStrAr As String() = {"tisíc", "tisíce", "tisíc", ""} ' all, 2-4, 1, 0
            Dim millionsToAddStrAr As String() = {"milionů", "miliony", "milion", ""} ' all, 2-4, 1, 0
            Dim crownsToAddStrAr As String() = {"korun českých", "koruny české", "koruna česká", "korun českých"} ' all, 2-4, 1, 0

            Dim myWordsWithInsertsStrLt As New List(Of String)()
            myWordsWithInsertsStrLt.AddRange(myWordsStrAr)

            For i = 0 To 8   ' will cycle throught all positions 1st[0] to 9th[8]
                If i = 0 Or i = 3 Or i = 6 Then

                    Dim internalArray As String()
                    Dim internalPos As Integer

                    If i = 0 Then
                        internalArray = millionsToAddStrAr
                        internalPos = 3
                    ElseIf i = 3 Then
                        internalArray = thousandsToAddStrAr
                        internalPos = 7
                    ElseIf i = 6 Then
                        internalArray = crownsToAddStrAr
                        internalPos = 11
                    End If

                    If myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 0 Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(3))
                    ElseIf myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 1 Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(2))
                    ElseIf (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 2) Or (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 3) Or (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 4) Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(1))
                    Else
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(0))
                    End If

                End If
            Next


            ' CONCATENATE strings into one string if not empty
            Dim finalWordStrBld As New System.Text.StringBuilder

            For Each item As String In myWordsWithInsertsStrLt
                If item <> "" Then
                    finalWordStrBld.AppendFormat("{0} ", item)
                End If
            Next

            returnWordStr = finalWordStrBld.ToString

            If inputNumberRoundedRemainderDb = 0 Then
                ' end of the MAIN FUNCTION is the inputNumber is an integer
                Return returnWordStr
            Else
                ' if the inputNumber is NOT an integer, ADD HALIRE
                Dim myCentsDigitsIntAr(1) As String
                myCentsDigitsIntAr = {"0", "0"}

                ' halire digits into an array
                If inputNumberRoundedRemainderDb.ToString.Length = 4 Then
                    For i = 0 To 3
                        If i = 2 Then
                            myCentsDigitsIntAr(0) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        ElseIf i = 3 Then
                            myCentsDigitsIntAr(1) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        End If
                    Next i
                Else
                    For i = 0 To 2
                        If i = 2 Then
                            myCentsDigitsIntAr(0) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        End If
                    Next i
                End If

                ' match halire words
                Dim myCentsWordsStrAr As String() = {"", ""}

                For i = 0 To 1 ' will cycle throught all digits 1st[0] to 6th[5]

                    Dim insertPosition As Integer = i ' last 6th position, indexed from 0
                    Dim insertWord As String ' select word to be added		

                    If i = 0 Then
                        If myCentsDigitsIntAr(i) > 1 Then
                            insertWord = tensStrAr(myCentsDigitsIntAr(i) - 1)
                        Else
                            'insertWord = tensStrAr(myCentsDigitsIntAr(i) - 1)								
                        End If
                    ElseIf i = 1 Then ' in the 2nd position
                        If (myCentsDigitsIntAr(i) = 0 And myCentsDigitsIntAr(i - 1) = 1) Then
                            insertWord = tensStrAr(myCentsDigitsIntAr(1))
                        ElseIf myCentsDigitsIntAr(i) = 0 Then
                            '
                        ElseIf myCentsDigitsIntAr(i - 1) = 1 Then
                            insertWord = teensStrAr(myCentsDigitsIntAr(i) - 1)
                        Else
                            insertWord = singlesStrAr(myCentsDigitsIntAr(i) - 1)
                        End If
                    End If

                    myCentsWordsStrAr.SetValue(insertWord, insertPosition) ' assign word into the position
                    insertWord = "" ' reset insert word or it inserts hundreds if current teens doesn't return any value
                Next


                ' add all words for halire

                Dim centsToAddStrAr As String() = {"haléřů", "haléře", "haléř", "-možnost neexistuje-"} ' all, 2-4, 1, 0

                finalWordStrBld.Append("a ")

                If inputNumberRoundedRemainderDb = 0.01 Then ' handles the one exception for 0.01
                    myCentsWordsStrAr = {"", "jeden"}
                End If

                For Each item As String In myCentsWordsStrAr
                    If item <> "" Then
                        finalWordStrBld.AppendFormat("{0} ", item)
                    End If
                Next

                If myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 1 Then
                    finalWordStrBld.Append(centsToAddStrAr(2))
                ElseIf (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 2) Or (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 3) Or (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 4) Then
                    finalWordStrBld.Append(centsToAddStrAr(1))
                Else
                    finalWordStrBld.Append(centsToAddStrAr(0))
                End If

                ' RETURN the result

                returnWordStr = finalWordStrBld.ToString

                Return returnWordStr

            End If ' end ADD HALIRE

        End If ' ends MAIN FUNCTION

    End Function



    Shared Function convertNumberToText_noSQL(ByVal i As Double) As String ' shared so no need to instantiate, just call the function; public by default
        ' Funtion summary:
        ' accepts a number of 1-9 integers and any number of decimal places, e.g. 1000,12999999
        ' returns the number as (Czech) text, e.g. jeden tisíc korun českých a třináct haléřů

        Dim inputNumberDb As Double = i ' accept number from the input box, e.g. 1000

        Dim inputNumberIntegerPortionDb As Double = Fix(inputNumberDb) ' from 1234.9 -> 1234

        Dim inputNumberRoundedRemainderDb As Double = Math.Round(inputNumberDb Mod 1, 2) ' from 1234.129 -> 0.13	

        Dim returnWordStr As String ' return word, e.g. jeden tisíc korun českých

        ' CHECKS the inputNumber
        If inputNumberDb = 0 Then
            returnWordStr = "nula korun českých"
            Return returnWordStr
        ElseIf inputNumberDb > 999999999 Then
            returnWordStr = "příliš velké číslo"
            Return returnWordStr
        Else
            ' MAIN FUNCTION for the integer portion of the number (the remainder portion of the number is dealt with separately, as will not be used often)
            ' convert number to digits
            Dim inputDigitsIntAr As Integer() = Array.ConvertAll(Of Char, Integer)(inputNumberIntegerPortionDb.ToString.ToCharArray, Function(c As Char) Integer.Parse(c.ToString)) '/ {1,2,3,4}
            Array.Reverse(inputDigitsIntAr)

            ' initialize an empty array of the length of the maximum number e.g. 999,999,999
            Dim myDigitsIntAr As Integer() = {0, 0, 0, 0, 0, 0, 0, 0, 0}

            ' assign number digits to an empty array
            For j = 0 To inputDigitsIntAr.Length - 1
                myDigitsIntAr.SetValue(inputDigitsIntAr(j), j)
            Next
            Array.Reverse(myDigitsIntAr) '/ {0,0,1,2,3,4}

            ' initialize an empty array of the length of the maximum number e.g. 999,999
            Dim myWordsStrAr As String() = {"", "", "", "", "", "", "", "", ""}

            ' mapping
            'Dim dataset As New DataSet() ' so far not used
            'dataset = GetFullSqlData() ' so far not used

            Dim hundredsStrAr As String() = {"jedno sto", "dvě sta", "tři sta",
                        "čtyři sta", "pět set", "šest set", "sedm set", "osm set", "devět set"}
            Dim tensStrAr As String() = {"deset", "dvacet", "třicet",
                        "čtyřicet", "padesát", "šedesát", "sedmdesát", "osmdesát", "devadesát"}
            Dim teensStrAr As String() = {"jedenáct", "dvanáct", "třináct",
                        "čtrnáct", "patnáct", "šestnáct", "sedmnáct", "osmnáct", "devatenáct"}
            Dim singlesStrAr As String() = {"jedna", "dva", "tři",
                        "čtyři", "pět", "šest", "sedm", "osm", "devět"}

            ' MATCH digits to mapping / e.g. {" "," ","jeden","dve sta","tricet","ctyri"}		
            For i = 0 To myDigitsIntAr.Length - 1 ' will cycle throught all digits 1st[0] to 6th[5]

                If myDigitsIntAr(i) = 0 Then ' if digit in a position is zero, do nothing
                    'do nothing
                Else ' if digit in a position is anything other than zero, 
                    Dim insertPosition As Integer = i ' the last possible position is 8th, indexed from 0
                    Dim insertWord As String ' the word to be added to the emptry myWordsStrAr array

                    If i = 0 Or i = 3 Or i = 6 Then
                        insertWord = hundredsStrAr(myDigitsIntAr(i) - 1) ' find word from array's position x-1 as indexed from 0					
                    ElseIf i = 1 Or i = 4 Or i = 7 Then
                        If myDigitsIntAr(i) = 1 Then
                            'do nothing
                        Else
                            insertWord = tensStrAr(myDigitsIntAr(i) - 1)
                        End If
                    ElseIf i = 2 Or i = 5 Or i = 8 Then
                        If myDigitsIntAr(i - 1) = 1 Then
                            insertWord = teensStrAr(myDigitsIntAr(i) - 1)
                        Else
                            insertWord = singlesStrAr(myDigitsIntAr(i) - 1)
                        End If
                    End If

                    myWordsStrAr.SetValue(insertWord, insertPosition) ' assign word into the position
                    insertWord = "" ' reset insert word or it inserts hundreds if current teens doesn't return any value
                End If
            Next

            If inputNumberIntegerPortionDb = 2 Then ' handles the one exception for when the inputNumberDb = 2
                myWordsStrAr = {"", "", "", "", "", "", "", "", "dvě"}
            End If


            ' ADD tisic or tisice		
            Dim thousandsToAddStrAr As String() = {"tisíc", "tisíce", "tisíc", ""} ' all, 2-4, 1, 0
            Dim millionsToAddStrAr As String() = {"milionů", "miliony", "milion", ""} ' all, 2-4, 1, 0
            Dim crownsToAddStrAr As String() = {"korun českých", "koruny české", "koruna česká", "korun českých"} ' all, 2-4, 1, 0

            Dim myWordsWithInsertsStrLt As New List(Of String)()
            myWordsWithInsertsStrLt.AddRange(myWordsStrAr)

            For i = 0 To 8   ' will cycle throught all positions 1st[0] to 9th[8]
                If i = 0 Or i = 3 Or i = 6 Then

                    Dim internalArray As String()
                    Dim internalPos As Integer

                    If i = 0 Then
                        internalArray = millionsToAddStrAr
                        internalPos = 3
                    ElseIf i = 3 Then
                        internalArray = thousandsToAddStrAr
                        internalPos = 7
                    ElseIf i = 6 Then
                        internalArray = crownsToAddStrAr
                        internalPos = 11
                    End If

                    If myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 0 Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(3))
                    ElseIf myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 1 Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(2))
                    ElseIf (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 2) Or (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 3) Or (myDigitsIntAr(i) = 0 And myDigitsIntAr(i + 1) = 0 And myDigitsIntAr(i + 2) = 4) Then
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(1))
                    Else
                        myWordsWithInsertsStrLt.Insert(internalPos, internalArray(0))
                    End If

                End If
            Next


            ' CONCATENATE strings into one string if not empty
            Dim finalWordStrBld As New System.Text.StringBuilder

            For Each item As String In myWordsWithInsertsStrLt
                If item <> "" Then
                    finalWordStrBld.AppendFormat("{0} ", item)
                End If
            Next

            returnWordStr = finalWordStrBld.ToString

            If inputNumberRoundedRemainderDb = 0 Then
                ' end of the MAIN FUNCTION is the inputNumber is an integer
                Return returnWordStr
            Else
                ' if the inputNumber is NOT an integer, ADD HALIRE
                Dim myCentsDigitsIntAr(1) As String
                myCentsDigitsIntAr = {"0", "0"}

                ' halire digits into an array
                If inputNumberRoundedRemainderDb.ToString.Length = 4 Then
                    For i = 0 To 3
                        If i = 2 Then
                            myCentsDigitsIntAr(0) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        ElseIf i = 3 Then
                            myCentsDigitsIntAr(1) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        End If
                    Next i
                Else
                    For i = 0 To 2
                        If i = 2 Then
                            myCentsDigitsIntAr(0) = inputNumberRoundedRemainderDb.ToString.Substring(i, 1)
                        End If
                    Next i
                End If

                ' match halire words
                Dim myCentsWordsStrAr As String() = {"", ""}

                For i = 0 To 1 ' will cycle throught all digits 1st[0] to 6th[5]

                    Dim insertPosition As Integer = i ' last 6th position, indexed from 0
                    Dim insertWord As String ' select word to be added		

                    If i = 0 Then
                        If myCentsDigitsIntAr(i) > 1 Then
                            insertWord = tensStrAr(myCentsDigitsIntAr(i) - 1)
                        Else
                            'insertWord = tensStrAr(myCentsDigitsIntAr(i) - 1)								
                        End If
                    ElseIf i = 1 Then ' in the 2nd position
                        If (myCentsDigitsIntAr(i) = 0 And myCentsDigitsIntAr(i - 1) = 1) Then
                            insertWord = tensStrAr(myCentsDigitsIntAr(1))
                        ElseIf myCentsDigitsIntAr(i) = 0 Then
                            '
                        ElseIf myCentsDigitsIntAr(i - 1) = 1 Then
                            insertWord = teensStrAr(myCentsDigitsIntAr(i) - 1)
                        Else
                            insertWord = singlesStrAr(myCentsDigitsIntAr(i) - 1)
                        End If
                    End If

                    myCentsWordsStrAr.SetValue(insertWord, insertPosition) ' assign word into the position
                    insertWord = "" ' reset insert word or it inserts hundreds if current teens doesn't return any value
                Next


                ' add all words for halire

                Dim centsToAddStrAr As String() = {"haléřů", "haléře", "haléř", "-možnost neexistuje-"} ' all, 2-4, 1, 0

                finalWordStrBld.Append("a ")

                If inputNumberRoundedRemainderDb = 0.01 Then ' handles the one exception for 0.01
                    myCentsWordsStrAr = {"", "jeden"}
                End If

                For Each item As String In myCentsWordsStrAr
                    If item <> "" Then
                        finalWordStrBld.AppendFormat("{0} ", item)
                    End If
                Next

                If myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 1 Then
                    finalWordStrBld.Append(centsToAddStrAr(2))
                ElseIf (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 2) Or (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 3) Or (myCentsDigitsIntAr(0) = 0 And myCentsDigitsIntAr(1) = 4) Then
                    finalWordStrBld.Append(centsToAddStrAr(1))
                Else
                    finalWordStrBld.Append(centsToAddStrAr(0))
                End If

                ' RETURN the result

                returnWordStr = finalWordStrBld.ToString

                Return returnWordStr

            End If ' end ADD HALIRE

        End If ' ends MAIN FUNCTION

    End Function

End Class

