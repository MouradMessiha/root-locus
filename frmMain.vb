Public Class frmMain

   Private Structure stcComplexNumber
      Public Real As Integer
      Public Imaginary As Integer
   End Structure

   Private marrPoles() As stcComplexNumber
   Private marrZeros() As stcComplexNumber

   Private mobjFormBitmap As Bitmap
   Private mobjBitmapGraphics As Graphics
   Private mintFormWidth As Integer
   Private mintFormHeight As Integer
   Private mintColorScheme As Integer = 0

   Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Static blnDoneOnce As Boolean = False

        If Not blnDoneOnce Then
            mintFormWidth = Me.Width
            mintFormHeight = Me.Height
            mobjFormBitmap = New Bitmap(mintFormWidth, mintFormHeight, Me.CreateGraphics())
            mobjBitmapGraphics = Graphics.FromImage(mobjFormBitmap)
            mobjBitmapGraphics.FillRectangle(Brushes.White, 0, 0, mintFormWidth, mintFormHeight)
            mobjBitmapGraphics.DrawLine(Pens.Black, 0, CType(mintFormHeight / 2, Int32), CType(mintFormWidth, Int32), CType(mintFormHeight / 2, Int32))
            mobjBitmapGraphics.DrawLine(Pens.Black, CType(mintFormWidth / 2, Int32), 0, CType(mintFormWidth / 2, Int32), CType(mintFormHeight, Int32))
            blnDoneOnce = True
        End If

   End Sub

   Private Sub frmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

      e.Graphics.DrawImage(mobjFormBitmap, 0, 0)

   End Sub

   Private Sub frmMain_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick

      Dim blnRealNumber As Boolean

      Select Case e.Button
         Case Windows.Forms.MouseButtons.Left
            If Math.Abs(e.Y - (mintFormHeight / 2)) < 5 Then
               If MessageBox.Show("Is it a real number?", "Plot Poles", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                  blnRealNumber = True
               Else
                  blnRealNumber = False
               End If
            Else
               blnRealNumber = False
            End If

            If marrPoles Is Nothing Then
               ReDim marrPoles(0)
            Else
               ReDim Preserve marrPoles(marrPoles.GetLength(0))
            End If
            If blnRealNumber Then
               marrPoles(marrPoles.GetLength(0) - 1).Real = e.X
               marrPoles(marrPoles.GetLength(0) - 1).Imaginary = mintFormHeight / 2
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X - 5, CType((mintFormHeight / 2), Int32) - 5, e.X + 5, CType((mintFormHeight / 2), Int32) + 5)
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X + 5, CType((mintFormHeight / 2), Int32) - 5, e.X - 5, CType((mintFormHeight / 2), Int32) + 5)
            Else
               marrPoles(marrPoles.GetLength(0) - 1).Real = e.X
               marrPoles(marrPoles.GetLength(0) - 1).Imaginary = e.Y
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X - 5, e.Y - 5, e.X + 5, e.Y + 5)
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X + 5, e.Y - 5, e.X - 5, e.Y + 5)

               ReDim Preserve marrPoles(marrPoles.GetLength(0))
               marrPoles(marrPoles.GetLength(0) - 1).Real = e.X
               marrPoles(marrPoles.GetLength(0) - 1).Imaginary = mintFormHeight - e.Y
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X - 5, mintFormHeight - e.Y - 5, e.X + 5, mintFormHeight - e.Y + 5)
               mobjBitmapGraphics.DrawLine(Pens.Black, e.X + 5, mintFormHeight - e.Y - 5, e.X - 5, mintFormHeight - e.Y + 5)
            End If

         Case Windows.Forms.MouseButtons.Right
            If Math.Abs(e.Y - (mintFormHeight / 2)) < 5 Then
               If MessageBox.Show("Is it a real number?", "Plot Zeros", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                  blnRealNumber = True
               Else
                  blnRealNumber = False
               End If
            Else
               blnRealNumber = False
            End If

            If marrZeros Is Nothing Then
               ReDim marrZeros(0)
            Else
               ReDim Preserve marrZeros(marrZeros.GetLength(0))
            End If
            If blnRealNumber Then
               marrZeros(marrZeros.GetLength(0) - 1).Real = e.X
               marrZeros(marrZeros.GetLength(0) - 1).Imaginary = mintFormHeight / 2
               mobjBitmapGraphics.DrawEllipse(Pens.Black, e.X - 5, CType((mintFormHeight / 2), Int32) - 5, 10, 10)
            Else
               marrZeros(marrZeros.GetLength(0) - 1).Real = e.X
               marrZeros(marrZeros.GetLength(0) - 1).Imaginary = e.Y
               mobjBitmapGraphics.DrawEllipse(Pens.Black, e.X - 5, e.Y - 5, 10, 10)

               ReDim Preserve marrZeros(marrZeros.GetLength(0))
               marrZeros(marrZeros.GetLength(0) - 1).Real = e.X
               marrZeros(marrZeros.GetLength(0) - 1).Imaginary = mintFormHeight - e.Y
               mobjBitmapGraphics.DrawEllipse(Pens.Black, e.X - 5, mintFormHeight - e.Y - 5, 10, 10)
            End If
      End Select

      Me.Invalidate()

   End Sub

   Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

      If e.KeyCode = Keys.Return Then
         DrawRootLocus()
      ElseIf e.KeyCode = Keys.P Then
         DrawColorSchemeNumber()
      End If

   End Sub

   Private Sub DrawRootLocus()

      Dim objPen As Pen
      Dim intX As Int16
      Dim intY As Int16
      Dim dblAngle As Double
      Dim dblAngleIncrement As Double
      Dim dblDX As Double
      Dim dblDY As Double
      Dim intCounter As Integer
      Dim lngColor As Long
      Dim intRed As Integer
      Dim intGreen As Integer
      Dim intBlue As Integer

      objPen = New Pen(Color.FromArgb(0, 0, 0))

      With mobjBitmapGraphics
         For intY = 0 To mintFormHeight / 2
            For intX = 0 To mintFormWidth
               dblAngle = 0
               If Not marrZeros Is Nothing Then
                  For intCounter = 0 To marrZeros.GetLength(0) - 1
                     dblDX = CType(intX, Double) - marrZeros(intCounter).Real
                     dblDY = CType(intY, Double) - marrZeros(intCounter).Imaginary
                     If dblDX = 0 Then
                        If dblDY > 0 Then
                           dblAngleIncrement = Math.PI / 2
                        Else
                           dblAngleIncrement = -Math.PI / 2
                        End If
                     Else
                        dblAngleIncrement = Math.Atan(dblDY / dblDX)
                        If dblDX < 0 Then
                           dblAngleIncrement += Math.PI
                        End If
                     End If
                     dblAngle += dblAngleIncrement
                  Next
               End If

               If Not marrPoles Is Nothing Then
                  For intCounter = 0 To marrPoles.GetLength(0) - 1
                     dblDX = CType(intX, Double) - marrPoles(intCounter).Real
                     dblDY = CType(intY, Double) - marrPoles(intCounter).Imaginary
                     If dblDX = 0 Then
                        If dblDY > 0 Then
                           dblAngleIncrement = Math.PI / 2
                        Else
                           dblAngleIncrement = -Math.PI / 2
                        End If
                     Else
                        dblAngleIncrement = Math.Atan(dblDY / dblDX)
                        If dblDX < 0 Then
                           dblAngleIncrement += Math.PI
                        End If
                     End If
                     dblAngle -= dblAngleIncrement
                  Next
               End If

               dblAngle *= (45 / Math.PI)
               lngColor = Math.Round(dblAngle)
               While lngColor > 90
                  lngColor -= 90
               End While

               While lngColor < 0
                  lngColor += 90
               End While

               If lngColor <= 45 Then
                  intRed = 63
                  intGreen = Math.Round(90 * Math.Sqrt(Math.Sqrt(45 - lngColor)))
                  intBlue = intGreen
               Else
                  intRed = 63
                  intGreen = Math.Round(90 * Math.Sqrt(Math.Sqrt(lngColor - 45)))
                  intBlue = 63
               End If

               Select Case mintColorScheme Mod 12
                  Case 0
                     objPen.Color = Color.FromArgb(intRed, intGreen, intBlue)
                  Case 1
                     objPen.Color = Color.FromArgb(intRed, intBlue, intGreen)
                  Case 2
                     objPen.Color = Color.FromArgb(intGreen, intRed, intBlue)
                  Case 3
                     objPen.Color = Color.FromArgb(intGreen, intBlue, intRed)
                  Case 4
                     objPen.Color = Color.FromArgb(intBlue, intRed, intGreen)
                  Case 5
                     objPen.Color = Color.FromArgb(intBlue, intGreen, intRed)
                  Case 6
                     objPen.Color = Color.FromArgb(255 - intRed, 255 - intGreen, 255 - intBlue)
                  Case 7
                     objPen.Color = Color.FromArgb(255 - intRed, 255 - intBlue, 255 - intGreen)
                  Case 8
                     objPen.Color = Color.FromArgb(255 - intGreen, 255 - intRed, 255 - intBlue)
                  Case 9
                     objPen.Color = Color.FromArgb(255 - intGreen, 255 - intBlue, 255 - intRed)
                  Case 10
                     objPen.Color = Color.FromArgb(255 - intBlue, 255 - intRed, 255 - intGreen)
                  Case 11
                     objPen.Color = Color.FromArgb(255 - intBlue, 255 - intGreen, 255 - intRed)
               End Select

               .DrawRectangle(objPen, intX, intY, 1, 1)
               .DrawRectangle(objPen, intX, mintFormHeight - intY, 1, 1)
            Next
         Next
      End With

      Me.Invalidate()

   End Sub

   Private Sub DrawColorSchemeNumber()

      Dim objFont As Font

      mintColorScheme += 1
      objFont = New Font("MS Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

      With mobjBitmapGraphics
         .FillRectangle(Brushes.White, 0, 0, 20, 20)
         .DrawString(mintColorScheme Mod 12, objFont, Brushes.Black, 0, 0)
      End With

      Me.Invalidate()

   End Sub

End Class

