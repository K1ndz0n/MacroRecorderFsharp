module funkcyjne_projekt.Main

open System.Collections.Generic
open System.Collections.ObjectModel
open System.Diagnostics
open System.IO
open Newtonsoft.Json
open Program
open Classes
open Structures
open System
open System.Windows.Forms
open Microsoft.FSharp.Core

           
let saveToFile (path: string, macros: Macro ObservableCollection) =
    let settings = JsonSerializerSettings(TypeNameHandling = TypeNameHandling.All)
    let json = JsonConvert.SerializeObject(macros, settings)
    File.WriteAllText(path, json)
    printfn "zapisano"
        
// Wczytanie z pliku
let loadFromFile (path: string) =
    try
        // Sprawdzenie, czy plik istnieje i nie jest pusty
        if File.Exists(path) && (new FileInfo(path)).Length > 0L then
            let json = File.ReadAllText(path)
            let settings = JsonSerializerSettings(TypeNameHandling = TypeNameHandling.All)
            JsonConvert.DeserializeObject<Macro ObservableCollection>(json, settings)
        else
            // Jeśli plik jest pusty lub nie istnieje, to zwróć pustą listę
            ObservableCollection<Macro>()
    with
    | ex -> 
        printfn "Błąd podczas wczytywania pliku: %s" ex.Message
        File.WriteAllText(path, "")
        ObservableCollection<Macro>()
   

        
let createEventDialog(eventType: string, macroEvents: ObservableCollection<MacroEvent>) =
    let dialog = new Form(Text = "Dodaj zdarzenie: " + eventType, Width = 300, Height = 260)
    
    match eventType with
    | "Klawisz systemowy" ->
        let comboBox = new ComboBox(Top = 30, Left = 20, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList)
        for kvp in sysKeys do
        comboBox.Items.Add(kvp.Value) |> ignore
        
        let labelComboBox = new Label(Text = "Wybierz klawisz:", AutoSize = true, Top = 10, Left = 20)
        dialog.Controls.Add(labelComboBox)
        dialog.Controls.Add(comboBox)
        
        let radioButtonPressed = new RadioButton(Text = "Wciśnięty", Top = 80, Left = 20, Checked = true)
        let radioButtonReleased = new RadioButton(Text = "Puszczony", Top = 110, Left = 20)

        let labelRadio = new Label(Text = "Stan klawisza:", AutoSize = true, Top = 60, Left = 20)
        dialog.Controls.Add(labelRadio)
        dialog.Controls.Add(radioButtonPressed)
        dialog.Controls.Add(radioButtonReleased)
        
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let selectedKey = comboBox.SelectedItem
            if selectedKey = null then
                MessageBox.Show("Nie wybrano przycisku!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let state = if radioButtonPressed.Checked then "Wciśnięty" else "Puszczony"
                let isPressed = state = "Wciśnięty"
                let mutable vkCode: int = -1
                // szukamy vkCode klawisza
                let found = 
                    sysKeys 
                    |> Seq.tryFind (fun kvp -> kvp.Value = string selectedKey)

                match found with
                | Some kvp ->
                    vkCode <- kvp.Key
                
                macroEvents.Add(SysKeyPressEvent((string)selectedKey, vkCode, isPressed))
                
                MessageBox.Show("Dodano klawisz systemowy: " + (string)selectedKey + ", stan: " + state) |> ignore
                dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)
        
    | "Ciąg znaków" ->
        let label = new Label(Top = 30, Left = 60, Text = "Wpisz ciąg znaków")
        let textBox = new TextBox(Top = 30, Left = 20, Width = 200, Dock = DockStyle.Top)
        dialog.Controls.Add(label)
        dialog.Controls.Add(textBox)
        
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let text: string = textBox.Text
            if text = null then
                MessageBox.Show("Nie wpisano ciągu znaków!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let mutable result: string = ""
                for c in text do
                    if (int)c >= 97 && (int)c <= 122 then
                        result <- result + (string)((char)((int)c - 32))
                    else if (int)c >= 65 && (int)c <= 90 then
                        result <- result + (string)c
                    else
                        result <- result + refactorChar((string)c)
                
                macroEvents.Add(KeyStroke(result))
                
                MessageBox.Show("Dodano ciąg znaków: " + result) |> ignore
                dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)
    
    | "Kliknięcie myszy" ->
        let labelX = new Label(Top = 5, Left = 20, Width = 30, Text = "X")
        let textBoxX = new TextBox(Top = 5, Left = 70, Width = 200)
        let labelY = new Label(Top = 35, Left = 20, Width = 30, Text = "Y")
        let textBoxY = new TextBox(Top = 35, Left = 70, Width = 200)
        
        let radioButtonLeft = new RadioButton(Text = "Lewy", Top = 80, Left = 20, Checked = true)
        let radioButtonRight = new RadioButton(Text = "Prawy", Top = 110, Left = 20)
        let radioButtonMid = new RadioButton(Text = "Środkowy", Top = 140, Left = 20)
        
        dialog.Controls.Add(labelX)
        dialog.Controls.Add(textBoxX)
        dialog.Controls.Add(labelY)
        dialog.Controls.Add(textBoxY)
        
        let labelRadio = new Label(Text = "Przycisk:", AutoSize = true, Top = 60, Left = 20)
        dialog.Controls.Add(labelRadio)
        dialog.Controls.Add(radioButtonLeft)
        dialog.Controls.Add(radioButtonRight)
        dialog.Controls.Add(radioButtonMid)
            
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let mutable state = null
            if radioButtonLeft.Checked then state <- "Lewy"
            else if radioButtonRight.Checked then state <- "Prawy"
            else state <- "Środkowy"
            
            let xpos = textBoxX.Text
            let ypos = textBoxY.Text
            if xpos = null || ypos = null || canParseToInt(xpos) = false || canParseToInt(ypos) = false then
                MessageBox.Show("Wpisz poprawne współrzędne myszy!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let pressedButton = if state = "Lewy" then "Left" else if state = "Prawy" then "Right" else "Middle"
                macroEvents.Add(MousePressEvent(pressedButton, (int)xpos, (int)ypos))
                MessageBox.Show("Dodano kliknięcie myszki: " + state + ", X: " + xpos + ", Y: " + ypos) |> ignore
                dialog.Close()           
        )
        dialog.Controls.Add(okButton)
        
        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)   
    
    | "Kółko myszy" ->
        let radioButtonUp = new RadioButton(Text = "W górę", Top = 80, Left = 20, Checked = true)
        let radioButtonDown = new RadioButton(Text = "W dół", Top = 110, Left = 20)
        dialog.Controls.Add(radioButtonUp)
        dialog.Controls.Add(radioButtonDown)
            
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let mutable state = null
            let mutable delta = -1
            if radioButtonUp.Checked then
                state <- "W górę"
                delta <- 120
            else if radioButtonDown.Checked then
                state <- "W dół"
                delta <- -120

            macroEvents.Add(MouseScrollEvent(delta))
            MessageBox.Show("Dodano ruszenie kółkiem: " + state) |> ignore
            dialog.Close()           
        )
        dialog.Controls.Add(okButton)
        
        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)
    
    | "Oczekiwanie" ->
        let label = new Label(Top = 30, Left = 60, Text = "Wpisz czas oczekiwania (ms)")
        let textBox = new TextBox(Top = 30, Left = 20, Width = 200, Dock = DockStyle.Top)
        dialog.Controls.Add(label)
        dialog.Controls.Add(textBox)
        
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let text: string = textBox.Text
            if text = null || canParseToInt(text) = false then
                MessageBox.Show("Wpisz poprawną wartość!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                if (int)text < 0 then
                    MessageBox.Show("Wpisz poprawną wartość!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                else                   
                    macroEvents.Add(WaitEvent((int)text))
                    MessageBox.Show("Dodano oczekiwanie: " + text) |> ignore
                    dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)

    dialog.ShowDialog() |> ignore


let editEvent(event: MacroEvent, index: int, macroEvents: ObservableCollection<MacroEvent>) =
    
    let dialog = new Form(Text = "Edytuj zdarzenie: ", Width = 300, Height = 260)
    
    match event with
    | :? SysKeyPressEvent ->
        let comboBox = new ComboBox(Top = 30, Left = 20, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList)
        for kvp in sysKeys do
        comboBox.Items.Add(kvp.Value) |> ignore
                
        let labelComboBox = new Label(Text = "Wybierz klawisz:", AutoSize = true, Top = 10, Left = 20)
        dialog.Controls.Add(labelComboBox)
        dialog.Controls.Add(comboBox)
        
        comboBox.SelectedItem <- (event :?> SysKeyPressEvent).Name
        
        
        let radioButtonPressed = new RadioButton(Text = "Wciśnięty", Top = 80, Left = 20)
        let radioButtonReleased = new RadioButton(Text = "Puszczony", Top = 110, Left = 20)
        if (event :?> SysKeyPressEvent).IsDown = true then
            radioButtonPressed.Checked <- true
        else
            radioButtonReleased.Checked <- true


        let labelRadio = new Label(Text = "Stan klawisza:", AutoSize = true, Top = 60, Left = 20)
        dialog.Controls.Add(labelRadio)
        dialog.Controls.Add(radioButtonPressed)
        dialog.Controls.Add(radioButtonReleased)
        

        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let selectedKey = comboBox.SelectedItem
            if selectedKey = null then
                MessageBox.Show("Nie wybrano przycisku!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let state = if radioButtonPressed.Checked then "Wciśnięty" else "Puszczony"
                let isPressed = state = "Wciśnięty"
                let mutable vkCode: int = -1
                // szukamy vkCode klawisza
                let found = 
                    sysKeys 
                    |> Seq.tryFind (fun kvp -> kvp.Value = string selectedKey)

                match found with
                | Some kvp ->
                    vkCode <- kvp.Key
                
                (macroEvents.[index] :?> SysKeyPressEvent).Name <- (string)selectedKey
                (macroEvents.[index] :?> SysKeyPressEvent).VkCode <- vkCode
                (macroEvents.[index] :?> SysKeyPressEvent).IsDown <- isPressed
                               
                MessageBox.Show("Nowa wartość: " + (string)selectedKey + ", stan: " + state) |> ignore
                dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)
        
    | :? KeyStroke ->
        let label = new Label(Top = 30, Left = 60, Text = "Wpisz ciąg znaków")
        let textBox = new TextBox(Top = 30, Left = 20, Width = 200, Dock = DockStyle.Top)
        dialog.Controls.Add(label)
        dialog.Controls.Add(textBox)
        
        textBox.Text <- (event :?> KeyStroke).Value
        
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let text: string = textBox.Text
            if text = null then
                MessageBox.Show("Nie wpisano ciągu znaków!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let mutable result: string = ""
                for c in text do
                    if (int)c >= 97 && (int)c <= 122 then
                        result <- result + (string)((char)((int)c - 32))
                    else if (int)c >= 65 && (int)c <= 90 then
                        result <- result + (string)c
                    else
                        result <- result + refactorChar((string)c)
                
                (macroEvents.[index] :?> KeyStroke).Value <- result
                
                MessageBox.Show("Nowa wartość: " + result) |> ignore
                dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)
    
    | :? MousePressEvent ->
        let labelX = new Label(Top = 5, Left = 20, Width = 30, Text = "X")
        let textBoxX = new TextBox(Top = 5, Left = 70, Width = 200)
        let labelY = new Label(Top = 35, Left = 20, Width = 30, Text = "Y")
        let textBoxY = new TextBox(Top = 35, Left = 70, Width = 200)
        
        let radioButtonLeft = new RadioButton(Text = "Lewy", Top = 80, Left = 20)
        let radioButtonRight = new RadioButton(Text = "Prawy", Top = 110, Left = 20)
        let radioButtonMid = new RadioButton(Text = "Środkowy", Top = 140, Left = 20)
        
        if (event :?> MousePressEvent).Button = "Left" then
            radioButtonLeft.Checked <- true
        else if (event :?> MousePressEvent).Button = "Right" then
            radioButtonRight.Checked <- true
        else
            radioButtonMid.Checked <- true
        
               
        dialog.Controls.Add(labelX)
        dialog.Controls.Add(textBoxX)
        dialog.Controls.Add(labelY)
        dialog.Controls.Add(textBoxY)
        
        textBoxX.Text <- (string)(event :?> MousePressEvent).PosX
        textBoxY.Text <- (string)(event :?> MousePressEvent).PosY
        
        let labelRadio = new Label(Text = "Przycisk:", AutoSize = true, Top = 60, Left = 20)
        dialog.Controls.Add(labelRadio)
        dialog.Controls.Add(radioButtonLeft)
        dialog.Controls.Add(radioButtonRight)
        dialog.Controls.Add(radioButtonMid)
            
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let mutable state = null
            if radioButtonLeft.Checked then state <- "Lewy"
            else if radioButtonRight.Checked then state <- "Prawy"
            else state <- "Środkowy"
            
            let xpos = textBoxX.Text
            let ypos = textBoxY.Text
            if xpos = null || ypos = null || canParseToInt(xpos) = false || canParseToInt(ypos) = false then
                MessageBox.Show("Wpisz poprawne współrzędne myszy!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let pressedButton = if state = "Lewy" then "Left" else if state = "Prawy" then "Right" else "Middle"
                
                (macroEvents.[index] :?> MousePressEvent).Button <- pressedButton
                (macroEvents.[index] :?> MousePressEvent).PosX <- (int)xpos
                (macroEvents.[index] :?> MousePressEvent).PosY <- (int)ypos
                
                MessageBox.Show("Nowa wartość: " + state + ", X: " + xpos + ", Y: " + ypos) |> ignore
                dialog.Close()           
        )
        dialog.Controls.Add(okButton)
        
        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)   
    
    | :? MouseScrollEvent ->
        let radioButtonUp = new RadioButton(Text = "W górę", Top = 80, Left = 20)
        let radioButtonDown = new RadioButton(Text = "W dół", Top = 110, Left = 20)
        dialog.Controls.Add(radioButtonUp)
        dialog.Controls.Add(radioButtonDown)
        
        if (event :?> MouseScrollEvent).Delta = 120 then
            radioButtonUp.Checked <- true
        else
            radioButtonDown.Checked <- true
            
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let mutable state = null
            let mutable delta = -1
            if radioButtonUp.Checked then
                state <- "W górę"
                delta <- 120
            else if radioButtonDown.Checked then
                state <- "W dół"
                delta <- -120

            (macroEvents.[index] :?> MouseScrollEvent).Delta <- delta
            MessageBox.Show("Nowa wartość: " + state) |> ignore
            dialog.Close()           
        )
        dialog.Controls.Add(okButton)
        
        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)  
    
    | :? WaitEvent ->
        let label = new Label(Top = 30, Left = 60, Text = "Wpisz czas oczekiwania (ms)")
        let textBox = new TextBox(Top = 30, Left = 20, Width = 200, Dock = DockStyle.Top)
        dialog.Controls.Add(label)
        dialog.Controls.Add(textBox)
        
        textBox.Text <- (string)(event :?> WaitEvent).TimeAmount
        
        let okButton = new Button(Text = "OK", Top = 180, Left = 20, Width = 80)
        okButton.Click.Add(fun _ ->
            let text: string = textBox.Text
            if text = null || canParseToInt(text) = false then
                MessageBox.Show("Wpisz poprawną wartość!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                if (int)text < 0 then
                    MessageBox.Show("Wpisz poprawną wartość!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                else
                    (macroEvents.[index] :?> WaitEvent).TimeAmount <- (int)text
                    
                    MessageBox.Show("Nowa wartość: " + text) |> ignore
                    dialog.Close()            
        )
        dialog.Controls.Add(okButton)

        let cancelButton = new Button(Text = "Anuluj", Top = 180, Left = 120, Width = 80)
        cancelButton.Click.Add(fun _ -> dialog.Close())
        dialog.Controls.Add(cancelButton)

    dialog.ShowDialog() |> ignore
    
    
[<STAThread>]
let editRecordedMacro(macroEvents: ObservableCollection<MacroEvent>) =
    
    let tempBuf = new ObservableCollection<MacroEvent>(macroEvents |> Seq.map (fun ev -> ev.Clone())) // bufor pomocniczy zeby na poczatku zmiany wprowadzac na kopii a nie na oryginale
    
    let form = new Form(Text = "Edytuj makro", Width = 900, Height = 600)
    
    // Lista zdarzeń do wyboru, potrzebne do panelu nagrywania i do panelu edycji zapisanych makr
    let events = ["Klawisz systemowy"; "Ciąg znaków"; "Kliknięcie myszy"; "Kółko myszy"; "Oczekiwanie"]

    // Pasek narzędzi
    let toolStrip1 = new ToolStrip()
    let saveButton = new ToolStripButton(Text = "Zapisz zmiany")
    let addEventButton = new ToolStripDropDownButton(Text = "Dodaj zdarzenie")
    let editEventButton = new ToolStripButton(Text = "Edytuj zdarzenie")
    let removeEventButton = new ToolStripButton(Text = "Usuń zdarzenie")

    editEventButton.Enabled <- false
        
    // Dodawanie typów zdarzeń do guzika dodawania
    for eventType in events do
        let menuItem = new ToolStripMenuItem(Text = eventType)
        menuItem.Click.Add(fun _ -> createEventDialog(eventType, tempBuf)) // Otwarcie okna dialogowego przy kliknięciu
        addEventButton.DropDownItems.Add(menuItem) |> ignore
       
    // Lista zarejestrowanych zdarzeń
    let eventListBox = new ListBox(Dock = DockStyle.Fill)
    for i in tempBuf do
        eventListBox.Items.Add(i.Display())
    
    tempBuf.CollectionChanged.Add(fun _ ->           // aktualizacja listy przy zmianie
        eventListBox.Items.Clear()
        for ev in tempBuf do
            eventListBox.Items.Add(ev.Display()) |> ignore
            
        if tempBuf.Count = 0 then
            saveButton.Enabled <- false
            editEventButton.Enabled <- false
        else if tempBuf.Count > 0 then
            saveButton.Enabled <- true
    )
    
    eventListBox.AllowDrop <- true

    let mutable draggedItemIndex = -1

    // Obsługa rozpoczęcia przeciągania
    eventListBox.MouseDown.Add(fun e ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            editEventButton.Enabled <- true
            eventListBox.DoDragDrop(eventListBox.SelectedItem, DragDropEffects.Move) |> ignore
    )

    eventListBox.DragOver.Add(fun e ->
        e.Effect <- DragDropEffects.Move
    )

    // Obsługa upuszczania elementu
    eventListBox.DragDrop.Add(fun e ->
        if draggedItemIndex >= 0 then
            let point = eventListBox.PointToClient(System.Drawing.Point(e.X, e.Y))
            let targetIndex = eventListBox.IndexFromPoint(point)

            if targetIndex >= 0 && targetIndex <> draggedItemIndex then
                // Aktualizacja kolejności w tempBuf
                let draggedItem = tempBuf.[draggedItemIndex]
                tempBuf.RemoveAt(draggedItemIndex)
                tempBuf.Insert(targetIndex, draggedItem)

                eventListBox.Items.Clear()
                for ev in tempBuf do
                    eventListBox.Items.Add(ev.Display()) |> ignore
    )
   
    saveButton.Click.Add(fun _ ->
        // Tworzymy okienko dialogowe do wpisania nazwy
        let inputDialog = new Form(Text = "Zapisz", Size = System.Drawing.Size(300, 150))
        let label = new Label(Text = "Czy chcesz zapisać zmiany?", AutoSize = true, Top = 10, Left = 20)
        let confirmButton = new Button(Text = "OK", Dock = DockStyle.Bottom, Top = 30, Left = 40)
        let cancelButton = new Button(Text = "Anuluj", Dock = DockStyle.Bottom, Top = 30, Left = 80)

        inputDialog.Controls.Add(label)
        inputDialog.Controls.Add(confirmButton)
        inputDialog.Controls.Add(cancelButton)

        confirmButton.Click.Add(fun _ ->
                macroEvents.Clear()
                // kopiuje nową zawartość
                for i in tempBuf do
                    macroEvents.Add(i)
                    
                inputDialog.Close()
                form.Close()
        )
        // Obsługa przycisku Anuluj
        cancelButton.Click.Add(fun _ ->
            inputDialog.Close()
        )
        
        inputDialog.ShowDialog() |> ignore             
    )
    
    editEventButton.Click.Add(fun _ ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            editEvent(tempBuf.[draggedItemIndex], draggedItemIndex, tempBuf)
            
            // Aktualizacja ListBox
            eventListBox.Items.Clear()
            for ev in tempBuf do
                eventListBox.Items.Add(ev.Display()) |> ignore
    )
    
    removeEventButton.Click.Add(fun _ ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            tempBuf.RemoveAt(draggedItemIndex)
    )

    // Dodawanie przycisków do paska narzędzi
    toolStrip1.Items.Add(saveButton)  
    toolStrip1.Items.Add(addEventButton)
    toolStrip1.Items.Add(editEventButton)
    toolStrip1.Items.Add(removeEventButton)
    
    // Panel do układania elementów
    let panel1 = new Panel(Dock = DockStyle.Fill)
    panel1.Controls.Add(eventListBox)
    panel1.Controls.Add(toolStrip1)

    form.Controls.Add(panel1)
    form.ShowDialog() |> ignore

type MainForm(recordButton: ToolStripButton, stopRecordButton: ToolStripButton) as this =
    inherit Form()
    
    let recordHotKey = -100
    let stopRecordHotKey = -200
    let runRecordedMacro = -300
    let terminateTask = -400

    do
        // Rejestracja skrótu Ctrl+Shift+R
        if not (RegisterHotKey(this.Handle, recordHotKey, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.R)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        // Rejestracja skrótu Ctrl+Shift+S
        if not (RegisterHotKey(this.Handle, stopRecordHotKey, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.S)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        // Rejestracja skrótu Ctrl+Shift+U
        if not (RegisterHotKey(this.Handle, runRecordedMacro, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.U)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        // Rejestracja skrótu Ctrl+Shift+T
        if not (RegisterHotKey(this.Handle, terminateTask, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.T)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        for i = 0 to recordedMacros.Count - 1 do
            if recordedMacros.[i].KeyBind <> null then
                if not (RegisterHotKey(this.Handle, i, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 (char recordedMacros.[i].KeyBind))) then
                    MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
        
        this.FormClosing.Add(fun _ -> 
            // Wyrejestrowanie skrótu przy zamknięciu
            UnregisterHotKey(this.Handle, recordHotKey) |> ignore
            UnregisterHotKey(this.Handle, stopRecordHotKey) |> ignore
            UnregisterHotKey(this.Handle, runRecordedMacro) |> ignore
            UnregisterHotKey(this.Handle, terminateTask) |> ignore
            for i = 0 to recordedMacros.Count - 1 do
                if recordedMacros.[i].KeyBind <> null then
                    UnregisterHotKey(this.Handle, i) |> ignore
        )

    member this.unbindAll() =
        UnregisterHotKey(this.Handle, recordHotKey) |> ignore
        UnregisterHotKey(this.Handle, stopRecordHotKey) |> ignore
        UnregisterHotKey(this.Handle, runRecordedMacro) |> ignore
        for i = 0 to recordedMacros.Count - 1 do
            if recordedMacros.[i].KeyBind <> null then
                UnregisterHotKey(this.Handle, i) |> ignore
                
    member this.bindAll() =
        if not (RegisterHotKey(this.Handle, recordHotKey, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.R)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore

        if not (RegisterHotKey(this.Handle, stopRecordHotKey, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.S)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore

        if not (RegisterHotKey(this.Handle, runRecordedMacro, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.U)) then
            MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        for i = 0 to recordedMacros.Count - 1 do
            if recordedMacros.[i].KeyBind <> null then
                if not (RegisterHotKey(this.Handle, i, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 (char recordedMacros.[i].KeyBind))) then
                    MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
    
    
    override this.WndProc(m: Message byref) =
        base.WndProc(&m)
        if m.Msg = WM_HOTKEY && m.WParam.ToInt32() = terminateTask then
            isHalted <- true

        
        if m.Msg = WM_HOTKEY && m.WParam.ToInt32() = recordHotKey then
            isHotkeyStarted <- true
            recordButton.PerformClick()
            UnregisterHotKey(this.Handle, recordHotKey) |> ignore
            UnregisterHotKey(this.Handle, runRecordedMacro) |> ignore
            for i = 0 to recordedMacros.Count - 1 do
                if recordedMacros.[i].KeyBind <> null then
                    UnregisterHotKey(this.Handle, i) |> ignore
            
        if m.Msg = WM_HOTKEY && m.WParam.ToInt32() = stopRecordHotKey then
            isHotkeyStopped <- true
            stopRecordButton.PerformClick()
            if not (RegisterHotKey(this.Handle, recordHotKey, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.R)) then
                MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            if not (RegisterHotKey(this.Handle, runRecordedMacro, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 Keys.U)) then
                MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
                
            for i = 0 to recordedMacros.Count - 1 do
            if recordedMacros.[i].KeyBind <> null then
                if not (RegisterHotKey(this.Handle, i, uint32 MOD_CONTROL ||| uint32 MOD_SHIFT, uint32 (char recordedMacros.[i].KeyBind))) then
                    MessageBox.Show("Nie udało się zarejestrować skrótu klawiszowego.") |> ignore
            
        if m.Msg = WM_HOTKEY && m.WParam.ToInt32() = runRecordedMacro then           
            this.unbindAll()
            let playbackTask =                
                System.Threading.Tasks.Task.Run(fun () ->                   
                    for i = 0 to macroBuf.Count - 1 do
                        if isHalted = false then
                            macroBuf.[i].Playback()
                        
                )   
            ()
            isHalted <- false
            this.bindAll()
                        
        for i = 0 to recordedMacros.Count - 1 do
            if recordedMacros.[i].KeyBind <> null then
                if m.Msg = WM_HOTKEY && m.WParam.ToInt32() = i then
                    this.unbindAll()
                    let playbackTask = 
                        System.Threading.Tasks.Task.Run(fun () ->                           
                            recordedMacros.[i].Playback()                            
                        )   
                    ()
                    this.bindAll()
                              
                
[<STAThread>]
[<EntryPoint>]
let main argv =
    let currentModule = GetModuleHandle(null)
    // wczytywanie zapisanych makr
    recordedMacros.Clear()          
    let readBuf = loadFromFile("macros.json")
    for i in readBuf do
        recordedMacros.Add(i)
    
    Application.ApplicationExit.Add(fun _ ->
        if mouseHookId <> IntPtr.Zero then
            UnhookWindowsHookEx(mouseHookId) |> ignore
            
        if keyboardHookId <> IntPtr.Zero then
            UnhookWindowsHookEx(keyboardHookId) |> ignore
        
        saveToFile("macros.json", recordedMacros)
    )
    let mutable isRecording = false
    
         
    let tabControl = new TabControl(Dock = DockStyle.Fill)
    
    
    let events = ["Klawisz systemowy"; "Ciąg znaków"; "Kliknięcie myszy"; "Kółko myszy"; "Oczekiwanie"]
    
    // PANEL NAGRYWANIA
    let tab1 = new TabPage(Text = "Nagrywanie")
    
    let toolStrip1 = new ToolStrip()
    let recordButton = new ToolStripButton(Text = "Nagraj (Ctrl+Shift+R)")
    let stopButton = new ToolStripButton(Text = "Zatrzymaj (Ctrl+Shift+S)")
    let saveButton = new ToolStripButton(Text = "Zapisz")
    let addEventButton = new ToolStripDropDownButton(Text = "Dodaj zdarzenie")
    let editEventButton = new ToolStripButton(Text = "Edytuj zdarzenie")
    let playButton = new ToolStripButton(Text = "Uruchom (Ctrl+Shift+U)")
    let removeEventButton = new ToolStripButton(Text = "Usuń zdarzenie")
    let clearListButton = new ToolStripButton(Text = "Wyczyść listę")
    
    // Tworzenie głównego okna
    let form = new MainForm(recordButton, stopButton)
    form.Text <- "Nagrywanie makr"
    form.Width <- 900
    form.Height <- 600
       
    saveButton.Enabled <- false
    stopButton.Enabled <- false
    playButton.Enabled <- false
    editEventButton.Enabled <- false
    clearListButton.Enabled <- false
    
    
    // Dodawanie typów zdarzeń do guzika dodawania
    for eventType in events do
        let menuItem = new ToolStripMenuItem(Text = eventType)
        menuItem.Click.Add(fun _ -> createEventDialog(eventType, macroBuf)) // Otwarcie okna dialogowego przy kliknięciu
        addEventButton.DropDownItems.Add(menuItem) |> ignore
        
                
    // Lista zarejestrowanych zdarzeń
    let eventListBox = new ListBox(Dock = DockStyle.Fill)
    macroBuf.CollectionChanged.Add(fun _ ->           // aktualizacja listy przy zmianie
        eventListBox.Items.Clear()
        for ev in macroBuf do
            eventListBox.Items.Add(ev.Display()) |> ignore
            
        if macroBuf.Count = 0 then
            saveButton.Enabled <- false
            playButton.Enabled <- false
            editEventButton.Enabled <- false
            clearListButton.Enabled <- false
        else if macroBuf.Count > 0 && isRecording = false then
            playButton.Enabled <- true
            saveButton.Enabled <- true
            clearListButton.Enabled <- true
    )
    
    eventListBox.AllowDrop <- true

    let mutable draggedItemIndex = -1

    // Obsługa rozpoczęcia przeciągania
    eventListBox.MouseDown.Add(fun e ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            editEventButton.Enabled <- true
            eventListBox.DoDragDrop(eventListBox.SelectedItem, DragDropEffects.Move) |> ignore
    )

    // Obsługa przeciągania nad listą
    eventListBox.DragOver.Add(fun e ->
        e.Effect <- DragDropEffects.Move
    )

    // Obsługa upuszczania elementu
    eventListBox.DragDrop.Add(fun e ->
        if draggedItemIndex >= 0 then
            let point = eventListBox.PointToClient(System.Drawing.Point(e.X, e.Y))
            let targetIndex = eventListBox.IndexFromPoint(point)

            if targetIndex >= 0 && targetIndex <> draggedItemIndex then
                // Aktualizacja kolejności w macroBuf
                let draggedItem = macroBuf.[draggedItemIndex]
                macroBuf.RemoveAt(draggedItemIndex)
                macroBuf.Insert(targetIndex, draggedItem)

                // Aktualizacja ListBox
                eventListBox.Items.Clear()
                for ev in macroBuf do
                    eventListBox.Items.Add(ev.Display()) |> ignore
    )
   
    recordButton.Click.Add(fun _ ->
        if mouseHookId = IntPtr.Zero || keyboardHookId = IntPtr.Zero then
            mouseHookId <- SetWindowsHookEx(WH_MOUSE_LL, mouseHookProc, currentModule, 0u)
            keyboardHookId <- SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookProc, currentModule, 0u)
            stopwatch <- Stopwatch.StartNew()
            printfn "Nagrywanie rozpoczęte"
            stopButton.Enabled <- true
            recordButton.Enabled <- false
            clearListButton.Enabled <- false
            isRecording <- true
    )
    
    stopButton.Click.Add(fun _ ->
        if mouseHookId <> IntPtr.Zero then
            UnhookWindowsHookEx(mouseHookId) |> ignore
            mouseHookId <- IntPtr.Zero
        
        if keyboardHookId <> IntPtr.Zero then
            UnhookWindowsHookEx(keyboardHookId) |> ignore
            keyboardHookId <- IntPtr.Zero

        printfn "Nagrywanie zatrzymane"
        isRecording <- false
        stopButton.Enabled <- false
        recordButton.Enabled <- true
        removeStartHotkey()
        macroBuf.RemoveAt(macroBuf.Count - 1)
        if macroBuf.[macroBuf.Count - 1] :? WaitEvent then
            macroBuf.RemoveAt(macroBuf.Count - 1)
            
        if macroBuf.Count > 0 then
            saveButton.Enabled <- true
            playButton.Enabled <- true
    )
   
    saveButton.Click.Add(fun _ ->
        // Tworzymy okienko dialogowe do wpisania nazwy
        let inputDialog = new Form(Text = "Podaj nazwę makra", Size = System.Drawing.Size(300, 150))
        let textBox = new TextBox(Dock = DockStyle.Top)
        let confirmButton = new Button(Text = "OK", Dock = DockStyle.Bottom)
        let cancelButton = new Button(Text = "Anuluj", Dock = DockStyle.Bottom)
        let buttonsPanel = new FlowLayoutPanel(Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft)
        buttonsPanel.Controls.AddRange([| confirmButton; cancelButton |])

        inputDialog.Controls.AddRange([| textBox; buttonsPanel |])

        let mutable macroName: string = null

        // Obsługa przycisku OK
        confirmButton.Click.Add(fun _ ->
            if not (String.IsNullOrWhiteSpace textBox.Text) then
                macroName <- textBox.Text
                inputDialog.Close()
        )
        // Obsługa przycisku Anuluj
        cancelButton.Click.Add(fun _ ->
            inputDialog.Close()
        )

        inputDialog.ShowDialog() |> ignore
        
        if macroName <> null then        
            if recordedMacros |> Seq.exists (fun i -> i.Name = macroName) then
                MessageBox.Show("Nazwa makra już istnieje!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                form.unbindAll()
                recordedMacros.Add(Macro(macroName, ObservableCollection<_>(macroBuf)))
                form.bindAll()
                macroBuf.Clear()               
                MessageBox.Show("Zapisano makro", "Zapis", MessageBoxButtons.OK) |> ignore              
    )
        
    playButton.Click.Add(fun _ ->
        form.unbindAll()
        let playbackTask = 
            System.Threading.Tasks.Task.Run(fun () ->
                for i = 0 to macroBuf.Count - 1 do
                    if isHalted = false then
                        macroBuf.[i].Playback()
                isHalted <- false
            )
        ()
        form.bindAll()
    )
    
    editEventButton.Click.Add(fun _ ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            editEvent(macroBuf.[draggedItemIndex], draggedItemIndex, macroBuf)
            
            // Aktualizacja ListBox
            eventListBox.Items.Clear()
            for ev in macroBuf do
                eventListBox.Items.Add(ev.Display()) |> ignore
    )
    
    clearListButton.Click.Add(fun _ ->
        // Tworzymy okienko dialogowe do wpisania nazwy
        let inputDialog = new Form(Text = "Wyczyść", Size = System.Drawing.Size(300, 150))
        let label = new Label(Text = "Czy na pewno chcesz wyczyścić listę?", AutoSize = true, Top = 10, Left = 20)
        let confirmButton = new Button(Text = "OK", Dock = DockStyle.Bottom, Top = 30, Left = 40)
        let cancelButton = new Button(Text = "Anuluj", Dock = DockStyle.Bottom, Top = 30, Left = 80)

        inputDialog.Controls.Add(label)
        inputDialog.Controls.Add(confirmButton)
        inputDialog.Controls.Add(cancelButton)

        confirmButton.Click.Add(fun _ ->
            macroBuf.Clear()   
            inputDialog.Close()
        )
        // Obsługa przycisku Anuluj
        cancelButton.Click.Add(fun _ ->
            inputDialog.Close()
        )    
        inputDialog.ShowDialog() |> ignore      
    )
    
    removeEventButton.Click.Add(fun _ ->
        if eventListBox.SelectedItem <> null then
            draggedItemIndex <- eventListBox.SelectedIndex
            macroBuf.RemoveAt(draggedItemIndex)
    )

    // Dodawanie przycisków do paska narzędzi
    toolStrip1.Items.Add(recordButton)
    toolStrip1.Items.Add(stopButton)
    toolStrip1.Items.Add(playButton)
    toolStrip1.Items.Add(saveButton)
    toolStrip1.Items.Add(addEventButton)
    toolStrip1.Items.Add(editEventButton)
    toolStrip1.Items.Add(clearListButton)
    toolStrip1.Items.Add(removeEventButton)
    
    // Panel do układania elementów
    let panel1 = new Panel(Dock = DockStyle.Fill)
    panel1.Controls.Add(eventListBox)
    panel1.Controls.Add(toolStrip1)

    tab1.Controls.Add(panel1)
    tabControl.TabPages.Add(tab1)
    
    
    // PANEL ZARZADZANIA NAGRANYMI MAKRAMI
    let tab2 = new TabPage(Text = "Nagrane makra")
    // Pasek narzędzi 2
    let toolStrip2 = new ToolStrip()
    let playSavedButton =  new ToolStripButton(Text = "Uruchom")
    let editMacroButton = new ToolStripButton(Text = "Edytuj")
    let renameButton = new ToolStripButton(Text = "Zmień nazwę")
    let addBindingButton = new ToolStripButton(Text = "Przypisz skróty")
    let removeMacroButton = new ToolStripButton(Text = "Usuń")
    
    playSavedButton.Enabled <- false
    editMacroButton.Enabled <- false
    renameButton.Enabled <- false
    addBindingButton.Enabled <- false
    removeMacroButton.Enabled <- false
        
    let mutable selectedMacroIndex = -1
    // lista zapisanych makr
    let macroListBox = new ListBox(Dock = DockStyle.Fill)
    recordedMacros.CollectionChanged.Add(fun _ ->           // aktualizacja listy przy zmianie
        macroListBox.Items.Clear()
        for mac in recordedMacros do
            macroListBox.Items.Add(mac.Display()) |> ignore
            
        playSavedButton.Enabled <- false
        editMacroButton.Enabled <- false
        renameButton.Enabled <- false
        addBindingButton.Enabled <- false
        removeMacroButton.Enabled <- false
    )
    for i in recordedMacros do
        macroListBox.Items.Add(i.Display())

    // Obsługa rozpoczęcia przeciągania
    macroListBox.MouseDown.Add(fun e ->
        if macroListBox.SelectedItem <> null then
            selectedMacroIndex <- macroListBox.SelectedIndex
            playSavedButton.Enabled <- true
            editMacroButton.Enabled <- true
            renameButton.Enabled <- true
            addBindingButton.Enabled <- true
            removeMacroButton.Enabled <- true
    )
    
    playSavedButton.Click.Add(fun _ ->
    if macroListBox.SelectedItem <> null then
        selectedMacroIndex <- macroListBox.SelectedIndex
        form.unbindAll()
        let playbackTask = 
            System.Threading.Tasks.Task.Run(fun () ->
                recordedMacros.[selectedMacroIndex].Playback()
            )
        ()
        form.bindAll()
    )

    
    editMacroButton.Click.Add(fun _ ->
        if macroListBox.SelectedItem <> null then
            selectedMacroIndex <- macroListBox.SelectedIndex
            editRecordedMacro(recordedMacros.[selectedMacroIndex].Events)
    )
    
    
    renameButton.Click.Add(fun _ ->
        if macroListBox.SelectedItem <> null then
            selectedMacroIndex <- macroListBox.SelectedIndex
                
            let inputDialog = new Form(Text = "Podaj nową nazwę", Size = System.Drawing.Size(300, 150))
            let textBox = new TextBox(Dock = DockStyle.Top)
            textBox.Text <- recordedMacros.[selectedMacroIndex].Name
            let confirmButton = new Button(Text = "OK", Dock = DockStyle.Bottom)
            let cancelButton = new Button(Text = "Anuluj", Dock = DockStyle.Bottom)
            let buttonsPanel = new FlowLayoutPanel(Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft)
            buttonsPanel.Controls.AddRange([| confirmButton; cancelButton |])

            inputDialog.Controls.AddRange([| textBox; buttonsPanel |])

            let mutable macroName: string = null

            // Obsługa przycisku OK
            confirmButton.Click.Add(fun _ ->
                if not (String.IsNullOrWhiteSpace textBox.Text) then
                    macroName <- textBox.Text
                    inputDialog.Close()
            )
            // Obsługa przycisku Anuluj
            cancelButton.Click.Add(fun _ ->
                inputDialog.Close()
            )

            inputDialog.ShowDialog() |> ignore
            
            if macroName <> null then        
                if recordedMacros |> Seq.exists (fun i -> i.Name = macroName) then
                    MessageBox.Show("Nazwa makra już istnieje!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                else
                    recordedMacros.[selectedMacroIndex].Name <- macroName
                    macroListBox.Items.Clear()
                    for mac in recordedMacros do
                        macroListBox.Items.Add(mac.Display()) |> ignore
    )
    
    addBindingButton.Click.Add(fun _ ->
         if macroListBox.SelectedItem <> null then
            selectedMacroIndex <- macroListBox.SelectedIndex
            // Tworzymy okienko dialogowe do wpisania nazwy
            let inputDialog = new Form(Text = "Podaj klawisz", Size = System.Drawing.Size(300, 150))
            let textBox = new TextBox(Dock = DockStyle.Top)
            if recordedMacros.[selectedMacroIndex].KeyBind <> null then textBox.Text <- recordedMacros.[selectedMacroIndex].KeyBind
            let confirmButton = new Button(Text = "OK", Dock = DockStyle.Bottom)
            let cancelButton = new Button(Text = "Anuluj", Dock = DockStyle.Bottom)
            let buttonsPanel = new FlowLayoutPanel(Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft)
            buttonsPanel.Controls.AddRange([| confirmButton; cancelButton |])

            inputDialog.Controls.AddRange([| textBox; buttonsPanel |])

            let mutable bindKey: string = null

            // Obsługa przycisku OK
            confirmButton.Click.Add(fun _ ->
                bindKey <- textBox.Text
                if bindKey = null || bindKey = "" then
                    MessageBox.Show("Usunięto przypisanie", "Przypisanie", MessageBoxButtons.OK) |> ignore
                    inputDialog.Close()
                    form.unbindAll()
                    recordedMacros.[selectedMacroIndex].KeyBind <- null
                    form.bindAll()
                    macroListBox.Items.Clear()
                    for i in recordedMacros do
                        macroListBox.Items.Add(i.Display())
                    
                else if bindKey.Length > 1 then
                    MessageBox.Show("Niepoprawny klawisz!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                    
                else
                    bindKey <- refactorChar(bindKey)
                    if (int (char bindKey)) >= 97 && (int (char bindKey)) <= 122 then
                        bindKey <- (string (char (int (char bindKey) - 32)))
                    
                    let mutable isBinded = false
                    for i in recordedMacros do
                        if i.KeyBind = bindKey || bindKey = "R" || bindKey = "S" || bindKey = "U" then
                            MessageBox.Show("Zajęty lub niedozwolony klawisz!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                            isBinded <- true
                    if isBinded = false then
                        form.unbindAll()
                        recordedMacros.[selectedMacroIndex].KeyBind <- bindKey
                        form.bindAll()
                        macroListBox.Items.Clear()
                        for i in recordedMacros do
                            macroListBox.Items.Add(i.Display())
                        MessageBox.Show("Dodano przypisanie", "Przypisanie", MessageBoxButtons.OK) |> ignore
                        inputDialog.Close()
            )
            // Obsługa przycisku Anuluj
            cancelButton.Click.Add(fun _ ->
                inputDialog.Close()
            )
            inputDialog.ShowDialog() |> ignore
    )
    
    
    removeMacroButton.Click.Add(fun _ ->
        if macroListBox.SelectedItem <> null then
            selectedMacroIndex <- macroListBox.SelectedIndex
            form.unbindAll()
            recordedMacros.RemoveAt(selectedMacroIndex)
            form.bindAll()
            playSavedButton.Enabled <- false
            editMacroButton.Enabled <- false
            renameButton.Enabled <- false
            addBindingButton.Enabled <- false
            removeMacroButton.Enabled <- false
    )
        
    // Dodawanie przycisków do paska narzędzi   
    toolStrip2.Items.Add(playSavedButton)
    toolStrip2.Items.Add(editMacroButton)
    toolStrip2.Items.Add(renameButton)
    toolStrip2.Items.Add(addBindingButton)
    toolStrip2.Items.Add(removeMacroButton)
    
    
    // Panel do układania elementów
    let panel2 = new Panel(Dock = DockStyle.Fill)
    panel2.Controls.Add(macroListBox)
    panel2.Controls.Add(toolStrip2)
    
    tab2.Controls.Add(panel2)

    
    tabControl.TabPages.Add(tab2)

    // Dodanie kontrolki TabControl do głównej formy
    form.Controls.Add(tabControl)

    Application.Run(form)
    
         
    0