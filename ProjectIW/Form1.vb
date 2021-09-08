Public Class Form1
    'storing the dimensions of the form
    Dim width As Double
    Dim height As Double

    Private Const winnigcoins As Integer = 4
    Private Const bordWidth As Integer = 7
    Private Const bordHeight As Integer = 6
    Private Const bordPortionOfWindow As Double = 0.75
    Private Const spaceportion As Double = 1 / 20
    Private Const buttonheightratio As Double = (1 - (bordHeight + 1 + 1) * spaceportion) / (bordHeight + 1)
    Private Const buttonwidthratio As Double = (bordPortionOfWindow - (bordWidth + 1) * spaceportion) / bordWidth


    'elements used in the game
    Dim muntButton(bordWidth) As Button
    Dim spelbord(bordWidth - 1, bordHeight - 1) As PictureBox

    Dim Player1Name As TextBox
    Dim Player2Name As TextBox

    Dim startbutton As Button = New Button()
    Dim stopbutton As Button = New Button()

    Dim player1score As Label
    Dim player2score As Label

    Dim currentplayer As Label = New Label()
    Dim currentname As Label = New Label()
    Dim currentcolor As PictureBox = New PictureBox()

    Dim huidigeKleur As Color
    Dim player1turn As Boolean = True
    Dim gamestarted As Boolean = False


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'dimensions of the form 
        width = Me.Size.Width
        height = Me.Size.Height

        'dimensions of the buttons and coins
        Dim buttonwidth As Double = width * buttonwidthratio
        Dim buttonheight As Double = height * buttonheightratio
        Dim spaceheight As Double = height * spaceportion
        Dim spacewidth As Double = width * spaceportion

        'adding the buttons to the form
        For i As Integer = 0 To (bordWidth - 1)
            muntButton(i) = New Button()
            muntButton(i).Text = "Press"
            muntButton(i).Width = buttonwidth
            muntButton(i).Height = buttonheight
            muntButton(i).Left = (i + 1) * spacewidth + i * buttonwidth
            muntButton(i).Top = spaceheight
            muntButton(i).Tag = i
            AddHandler muntButton(i).Click, AddressOf Button_Click
            Me.Controls.Add(muntButton(i))
        Next

        'adding the coins to the form
        For i As Integer = spelbord.GetLowerBound(0) To spelbord.GetUpperBound(0)
            For j As Integer = spelbord.GetLowerBound(1) To spelbord.GetUpperBound(1)
                Dim munt As PictureBox = New PictureBox()
                munt.Width = buttonwidth
                munt.Height = buttonheight
                munt.BackColor = Color.White
                munt.Left = (i + 1) * spacewidth + i * buttonwidth
                munt.Top = (j + 2) * spaceheight + (j + 1) * buttonheight
                munt.Tag = Tuple.Create(i, j)
                spelbord(i, j) = munt
                Me.Controls.Add(spelbord(i, j))

            Next
        Next

        Dim startwidth As Double = width * bordPortionOfWindow
        Dim labelwidth As Double = width * (1 - bordPortionOfWindow - spaceportion)
        Dim Player As Label = New Label()
        Player.Text = "Player1"
        Player.Height = spaceheight
        Player.Width = labelwidth
        Player.Left = startwidth
        Player.Top = spaceheight
        Me.Controls.Add(Player)

        Player1Name = New TextBox()
        Player1Name.Height = spaceheight
        Player1Name.Width = labelwidth
        Player1Name.Left = startwidth
        Player1Name.Top = 2 * spaceheight
        Me.Controls.Add(Player1Name)

        Dim score As Label = New Label()
        score.Text = "score"
        score.Height = spaceheight
        score.Width = labelwidth
        score.Left = startwidth
        score.Top = 3 * spaceheight
        Me.Controls.Add(score)

        player1score = New Label()
        player1score.Height = spaceheight
        player1score.Width = labelwidth
        player1score.Left = startwidth
        player1score.Top = 4 * spaceheight
        player1score.BackColor = Color.White
        player1score.Text = "0"
        Me.Controls.Add(player1score)


        Dim Player2 As Label = New Label()
        Player2.Text = "Player2"
        Player2.Height = spaceheight
        Player2.Width = labelwidth
        Player2.Left = startwidth
        Player2.Top = 6 * spaceheight
        Me.Controls.Add(Player2)

        Player2Name = New TextBox()
        Player2Name.Height = spaceheight
        Player2Name.Width = labelwidth
        Player2Name.Left = startwidth
        Player2Name.Top = 7 * spaceheight
        Me.Controls.Add(Player2Name)

        Dim score2 As Label = New Label()
        score2.Text = "score"
        score2.Height = spaceheight
        score2.Width = labelwidth
        score2.Left = startwidth
        score2.Top = 8 * spaceheight
        Me.Controls.Add(score2)

        player2score = New Label()
        player2score.Height = spaceheight
        player2score.Width = labelwidth
        player2score.Left = startwidth
        player2score.Top = 9 * spaceheight
        player2score.BackColor = Color.White
        player2score.Text = 0
        Me.Controls.Add(player2score)

        startbutton.Text = "Start"
        startbutton.Height = spaceheight
        startbutton.Width = labelwidth / 2
        startbutton.Left = startwidth
        startbutton.Top = 10 * spaceheight
        AddHandler startbutton.Click, AddressOf Start_Button_Click
        Me.Controls.Add(startbutton)

        stopbutton.Text = "Stop"
        stopbutton.Height = spaceheight
        stopbutton.Width = labelwidth / 2
        stopbutton.Left = startwidth + labelwidth / 2
        stopbutton.Top = 10 * spaceheight
        AddHandler stopbutton.Click, AddressOf Stop_Button_Click
        Me.Controls.Add(stopbutton)


        currentplayer.Text = "Current Player"
        currentplayer.Height = spaceheight
        currentplayer.Width = labelwidth
        currentplayer.Left = startwidth
        currentplayer.Top = 12 * spaceheight
        Me.Controls.Add(currentplayer)

        currentcolor.Height = buttonheight
        currentcolor.Width = labelwidth / 2
        currentcolor.Left = startwidth
        currentcolor.Top = 13 * spaceheight
        currentcolor.BackColor = Color.Red
        Me.Controls.Add(currentcolor)

        currentname.Height = buttonheight
        currentname.Width = labelwidth / 2
        currentname.Left = startwidth + labelwidth / 2
        currentname.Top = 13 * spaceheight
        Me.Controls.Add(currentname)



    End Sub

    Private Sub Start_Button_Click(sender As Object, e As EventArgs)
        gamestarted = True
    End Sub

    Private Sub Stop_Button_Click(sender As Object, e As EventArgs)
        gamestarted = False
        player1score.Text = 0
        player2score.Text = 0
        resetBord()
    End Sub

    Public Function hasGameStarted() As Boolean
        Return gamestarted
    End Function

    Private Sub Button_Click(sender As Object, e As EventArgs)
        If (Not hasGameStarted()) Then
            Exit Sub
        End If
        Dim kleur As Color = getColor()
        Dim toBeColored As PictureBox = getToBeColoredCoin(sender)
        If Not IsNothing(toBeColored) Then
            toBeColored.BackColor = kleur
        End If
        If newcoloredhaswon(toBeColored) Then
            If player1turn Then
                MessageBox.Show(Player1Name.Text & " has won")
                player1score.Text = CInt(player1score.Text) + 1
            Else
                MessageBox.Show(Player2Name.Text & " has won")
                player2score.Text = CInt(player2score.Text) + 1
            End If ' bord should become white again and the color should also be reset
            resetBord()
            Exit Sub
        End If
        If bordisfull() Then
            resetBord()
            Exit Sub
        End If
        changePlayer()
        changeCurrentColorBox()
    End Sub

    'returns the color which has to be used
    Private Function getColor() As Color
        If player1turn Then
            Return Color.Red
        End If
        Return Color.Yellow
    End Function

    'resets the bord after one player has won
    Private Sub resetBord()
        For Each e As PictureBox In spelbord
            e.BackColor = Color.White
        Next
        player1turn = True
        changeCurrentColorBox()
    End Sub

    'checks whether the game has been fully colored
    Private Function bordisfull() As Boolean
        For Each e As PictureBox In spelbord
            If e.BackColor = Color.White Then
                Return False
            End If
        Next
        Return True
    End Function

    'changes the turn
    Private Sub changePlayer()
        If player1turn Then
            player1turn = False
        Else
            player1turn = True
        End If

    End Sub

    Private Sub changeCurrentColorBox()
        If player1turn Then
            currentcolor.BackColor = Color.Red
            currentname.Text = Player1Name.Text
        Else
            currentcolor.BackColor = Color.Yellow
            currentname.Text = Player2Name.Text
        End If
    End Sub

    'checks whether the new colored coin results in a victory
    Private Function newcoloredhaswon(coloredcoin As PictureBox) As Boolean
        If IsNothing(coloredcoin) Then
            Return False
        End If
        If verticalCheck(coloredcoin) Or horizontalCheck(coloredcoin) Or diagonalnegativecheck(coloredcoin) Or diagonalpositivecheck(coloredcoin) Then
            Return True
        End If
        Return False
    End Function

    Private Function horizontalCheck(coloredcoin As PictureBox) As Boolean
        Dim color As Color = coloredcoin.BackColor
        Dim vertical As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Dim leftcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim rightcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim numberofcolored As Integer = -1
        Do While spelbord(leftcolored, vertical).BackColor = color
            leftcolored -= 1
            If leftcolored = -1 Then
                Exit Do
            End If
        Loop
        Do While spelbord(rightcolored, vertical).BackColor = color
            rightcolored += 1
            If rightcolored = bordWidth Then
                Exit Do
            End If
        Loop
        numberofcolored = rightcolored - leftcolored - 1
        If numberofcolored >= winnigcoins Then
            Return True
        End If
        Return False

    End Function

    Private Function verticalCheck(coloredcoin As PictureBox) As Boolean
        Dim color As Color = coloredcoin.BackColor
        Dim horizontal As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim topcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Dim bottomcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Dim numberofcolored As Integer = -1
        Do While spelbord(horizontal, topcolored).BackColor = color
            topcolored -= 1
            If topcolored = -1 Then
                Exit Do
            End If
        Loop
        Do While spelbord(horizontal, bottomcolored).BackColor = color
            bottomcolored += 1
            If bottomcolored = bordHeight Then
                Exit Do
            End If
        Loop
        numberofcolored = bottomcolored - topcolored - 1
        If numberofcolored >= winnigcoins Then
            Return True
        End If
        Return False
    End Function

    Private Function diagonalpositivecheck(coloredcoin As PictureBox) As Boolean
        Dim color As Color = coloredcoin.BackColor
        Dim vertical As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Dim leftcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim rightcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim numberofcolored As Integer = -1
        Do While spelbord(leftcolored, vertical).BackColor = color
            leftcolored -= 1
            vertical += 1
            If leftcolored = -1 Then
                Exit Do
            End If
            If vertical = bordHeight Then
                Exit Do
            End If
        Loop
        vertical = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Do While spelbord(rightcolored, vertical).BackColor = color
            rightcolored += 1
            vertical -= 1
            If rightcolored = bordWidth Then
                Exit Do
            End If
            If vertical = -1 Then
                Exit Do
            End If
        Loop
        numberofcolored = rightcolored - leftcolored - 1
        If numberofcolored >= winnigcoins Then
            Return True
        End If
        Return False

    End Function

    Private Function diagonalnegativecheck(coloredcoin As PictureBox) As Boolean
        Dim color As Color = coloredcoin.BackColor
        Dim vertical As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Dim leftcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim rightcolored As Integer = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item1
        Dim numberofcolored As Integer = -1
        Do While spelbord(leftcolored, vertical).BackColor = color
            leftcolored -= 1
            vertical -= 1
            If leftcolored = -1 Then
                Exit Do
            End If
            If vertical = -1 Then
                Exit Do
            End If
        Loop
        vertical = CType(coloredcoin.Tag, Tuple(Of Integer, Integer)).Item2
        Do While spelbord(rightcolored, vertical).BackColor = color
            rightcolored += 1
            vertical += 1
            If rightcolored = bordWidth Then
                Exit Do
            End If
            If vertical = bordHeight Then
                Exit Do
            End If
        Loop
        numberofcolored = rightcolored - leftcolored - 1
        If numberofcolored >= winnigcoins Then
            Return True
        End If
        Return False
    End Function

    'return the coin which has to get colored and if the raw is full returns nothing 
    Private Function getToBeColoredCoin(pressedbutton As Button) As PictureBox
        Dim i As Integer = pressedbutton.Tag
        Dim j As Integer = bordHeight - 1
        Do While spelbord(i, j).BackColor <> Color.White
            j -= 1
            If j = -1 Then
                Exit Do
            End If
        Loop

        If j = -1 Then
            Return Nothing
        End If

        Return spelbord(i, j)

    End Function

End Class
