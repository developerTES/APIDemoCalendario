Imports System.IO
Imports System.Threading

Imports Google.Apis.Calendar.v3
Imports Google.Apis.Calendar.v3.Data
Imports Google.Apis.Calendar.v3.EventsResource
Imports Google.Apis.Services
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Util.Store



Module Module1

    'Dim scopes As IList(Of String) = New List(Of String)()
    Dim scopes As IList(Of String) = {"profile", "email", "CalendarService.Scope.Calendar"}
    Dim service As CalendarService
    Dim calendarId As String = "developer@englishschool.edu.co"
    Sub Main()


        Console.WriteLine("Calendar API Sample: List MyCalendar")
        Console.WriteLine("================================")

        Try
            Dim credential As UserCredential

            Using stream = New FileStream("client_secrets.json", FileMode.Open, FileAccess.Read)
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, {CalendarService.Scope.Calendar}, "user", CancellationToken.None, New FileDataStore("Calendar.MyEvents")).Result
            End Using

            Dim service = New CalendarService(New BaseClientService.Initializer() With {
            .HttpClientInitializer = credential,
            .ApplicationName = "Calendar API Sample"
        })


            Dim calendarList = service.CalendarList.List.ExecuteAsync()

            Dim Calendar = service.Calendars.Get(calendarId).Execute()
            Console.WriteLine("Calendar Name :")
            Console.WriteLine(Calendar.Summary)
            Console.WriteLine(Calendar.Description)


            '''''Metodos con calendarios
            '''
            ListarEventos(service)
            'crearEvento(service)





        Catch ex As AggregateException

            For Each e In ex.InnerExceptions
                Console.WriteLine("ERROR: " & e.Message)
            Next
        End Try

        Console.WriteLine("Press any key to continue...")
        Console.ReadKey()
    End Sub



    Private Sub ListarEventos(ByVal service As CalendarService)

        Try
            '// Define parameters of request.
            Dim ListRequest As ListRequest = service.Events.List(calendarId)

            ListRequest.TimeMin = DateTime.Now
            ListRequest.ShowDeleted = False
            ListRequest.SingleEvents = True
            ListRequest.MaxResults = 10
            ListRequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime

            '// List events.
            Dim Events As Events = ListRequest.Execute()
            Console.WriteLine("Upcoming events:")
            If Events.Items IsNot Nothing And Events.Items.Count > 0 Then
                Console.WriteLine("Si  hay eventos!")

                For Each ev As Data.Event In Events.Items
                    Dim hora As String = ev.Start.DateTime.ToString()
                    Console.WriteLine("{0} ({1})", ev.Summary, hora)
                Next






            Else
                Console.WriteLine("No upcoming events found.")
            End If
            Console.WriteLine("click for more .. ")
            Console.Read()
        Catch ex As Exception
            Console.WriteLine("Excepcion .. ")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
            Console.Read()
        End Try







    End Sub

    Private Sub crearEvento(ByVal service As CalendarService)

        Try
            Dim ev = New Data.Event()
            Dim start = New EventDateTime()
            start.DateTime = New DateTime(2023, 2, 5, 22, 0, 0)

            Dim end_Date = New EventDateTime()
            end_Date.DateTime = New DateTime(2023, 2, 5, 22, 30, 0)


            ev.Start = start
            ev.End = end_Date
            ev.Summary = "New Event"
            ev.Description = "Description..."
            Dim recurringEvent = service.Events.Insert(ev, calendarId).Execute()
            Console.WriteLine("Event created: %s\n", ev.HtmlLink)
        Catch ex As Exception
            Console.WriteLine("Exception ..! ")
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine(ex.Message)
            Console.Read()
        End Try


    End Sub


    Private Sub DisplayList(list As IList(Of CalendarListEntry))
        Console.WriteLine("Lists of calendars:")
        For Each item As CalendarListEntry In list
            Console.WriteLine(item.Summary & ". Location: " & item.Location & ", TimeZone: " & item.TimeZone)
        Next
    End Sub

    Private Sub DisplayFirstCalendarEvents(list As CalendarListEntry)
        Console.WriteLine(Environment.NewLine & "Maximum 5 first events from {0}:", list.Summary)
        Dim requeust As ListRequest = service.Events.List(list.Id)
        ' Set MaxResults and TimeMin with sample values
        requeust.MaxResults = 5
        requeust.TimeMin = New DateTime(2013, 10, 1, 20, 0, 0)
        ' Fetch the list of events
        For Each calendarEvent As Data.Event In requeust.Execute().Items
            Dim startDate As String = "Unspecified"
            If (Not calendarEvent.Start Is Nothing) Then
                If (Not calendarEvent.Start.Date Is Nothing) Then
                    startDate = calendarEvent.Start.Date.ToString()
                End If
            End If

            Console.WriteLine(calendarEvent.Summary & ". Start at: " & startDate)
        Next
    End Sub




End Module