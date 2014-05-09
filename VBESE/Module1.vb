
Module Module1

    Sub Main()
        Dim db As New Database()
        db.CreateInstance()
        db.CreateDatabase()

        Dim ev1 = New Item() With { _
            .Id = Guid.NewGuid(), _
            .Description = "Test", _
            .Price = 100.0, _
            .StartTIme = DateTime.Now _
        }
        db.AddEvent(ev1)
        Dim events = db.GetAllEvents()
        For Each ev As Item In events
            Console.WriteLine(ev.ToString())
        Next

        For i As Integer = 0 To 100
            Dim ev = New Item() With { _
                .Id = Guid.NewGuid(), _
                .Description = "Test" + i.ToString(), _
                .Price = 0.0 + Convert.ToDouble(i), _
                .StartTIme = DateTime.Now _
            }
            db.AddEvent(ev)
        Next
        Dim events2 = db.GetEventsForPriceRange(0, 61)
        For Each ev As Item In events2
            Console.WriteLine(ev.ToString())
        Next
        Console.ReadLine()

    End Sub

End Module


