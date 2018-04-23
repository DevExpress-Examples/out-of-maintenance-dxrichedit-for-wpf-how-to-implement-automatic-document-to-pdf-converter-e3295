Imports Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports System.Windows

Imports DevExpress.XtraRichEdit

Namespace DocumentServer_PrintToPDF
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Private counter As PerformanceCounter
		Private executing As Boolean

		Public Sub New()
			InitializeComponent()

			tbPath.Text = Application.Current.StartupUri.AbsolutePath

			Dim procName As String = Process.GetCurrentProcess().ProcessName
			counter = New PerformanceCounter("Process", "Working Set - Private", procName)
			ShowMemoryUsage()
		End Sub

		Private Function PrintToPDF(ByVal server As RichEditDocumentServer, ByVal filePath As String) As String
			Try
				server.LoadDocument(filePath)
			Catch ex As Exception
				server.CreateNewDocument()
				Return String.Format("{0:T} Error:{1} -> {2}", DateTime.Now, ex.Message, filePath) & Environment.NewLine
			End Try
			Dim outFileName As String = Path.ChangeExtension(filePath, "pdf")
			Dim fsOut As FileStream = File.Open(outFileName, FileMode.Create)
			server.ExportToPdf(fsOut)
			fsOut.Close()
			Return String.Format("{0:T} Done-> {1}", DateTime.Now, outFileName) & Environment.NewLine
		End Function

		Private Sub ConvertFiles(ByVal path As String)
			If (Not Directory.Exists(path)) Then
				Return
			End If

			Dim files() As String = System.IO.Directory.GetFiles(path, "*.doc?", System.IO.SearchOption.AllDirectories)
			InitProgress(files.Length)
			Using server As New RichEditDocumentServer()
				Dim count As Integer = files.Length
				For i As Integer = 0 To count - 1
					Dim progress As String = PrintToPDF(server, files(i))
					AppendProgress(progress, i + 1)
					If (Not executing) Then
						Return
					End If
				Next i
			End Using
		End Sub

		Private Sub btnConvert_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If (Not executing) Then
				StartConversion()
			Else
				FinishConversion()
			End If
		End Sub

		Private Sub StartConversion()
			executing = True
			btnConvert.Content = "Stop!"
			tbPath.IsReadOnly = True
			pnlProgress.Visibility = System.Windows.Visibility.Visible
			Dim worker As New Thread(AddressOf BackgroundWorker)
			worker.Start(tbPath.Text)
		End Sub
		Private Sub FinishConversion()
			pnlProgress.Visibility = System.Windows.Visibility.Collapsed
			executing = False
			tbPath.IsReadOnly = False
			btnConvert.Content = "Start!"
		End Sub
		Private Sub BackgroundWorker(ByVal parameter As Object)
			Dim path As String = TryCast(parameter, String)
			If String.IsNullOrEmpty(path) Then
				Return
			End If
			ConvertFiles(path)
			ShowMemoryUsage()

			Dim action As Action = Function() AnonymousMethod1()
			Dispatcher.BeginInvoke(action)
		End Sub
		
		Private Function AnonymousMethod1() As Boolean
			FinishConversion()
			Return True
		End Function

		Private Sub InitProgress(ByVal fileCount As Integer)
			Dim action As Action = Function() AnonymousMethod2(fileCount)
			Me.Dispatcher.Invoke(action)
		End Sub
		
		Private Function AnonymousMethod2(ByVal fileCount As Integer) As Boolean
			edtProgress.Minimum = 0
			edtProgress.Maximum = fileCount
			edtProgress.EditValue = 0
			Return True
		End Function
		Private Sub AppendProgress(ByVal displayText As String, ByVal fileIndex As Integer)
			Dim action As Action = Function() AnonymousMethod3(displayText, fileIndex)
			Me.Dispatcher.Invoke(action)
		End Sub
		
		Private Function AnonymousMethod3(ByVal displayText As String, ByVal fileIndex As Integer) As Boolean
			tbLog.Text += displayText
			LogScrollViewer.ScrollToVerticalOffset(LogScrollViewer.Height)
			edtProgress.EditValue = fileIndex
			Return True
		End Function
		Private Sub ShowMemoryUsage()
			Dim action As Action = Function() AnonymousMethod4()
			Me.Dispatcher.Invoke(action)
		End Sub
		
		Private Function AnonymousMethod4() As Boolean
			lblMemoryUsage.Text = String.Format("Memory usage: {0:N0} K", counter.RawValue / 1024)
			Return True
		End Function
		Private Sub Window_Closed(ByVal sender As Object, ByVal e As EventArgs)
			executing = False
		End Sub
	End Class
End Namespace
