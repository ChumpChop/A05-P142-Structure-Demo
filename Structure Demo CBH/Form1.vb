Public Class Form1
    'Form level members
    'Private objCustomers As New ArrayList
    Private objCustomers As New CustomerCollection

    Public ReadOnly Property SelectedCustomer() As Customer
        Get
            If lstCustomers.SelectedIndex <> -1 Then
                'Return the selected customer
                'Return CType(objCustomers(lstCustomers.SelectedIndex), Customer)
                Return objCustomers(lstCustomers.SelectedIndex)
            End If
        End Get
    End Property

    Public Sub CreateCustomer(ByVal firstName As String, ByVal lastName As String, ByVal email As String)

        'Declare a Customer object
        Dim objNewCustomer As Customer

        'Create the new customer
        objNewCustomer.FirstName = firstName
        objNewCustomer.LastName = lastName
        objNewCustomer.Email = email

        'Add the new customer to the list
        objCustomers.Add(objNewCustomer)

        'Add the new customer to the ListBox control
        lstCustomers.Items.Add(objNewCustomer)
    End Sub

    Public Sub DisplayCustomer(ByVal customer As Customer)
        'Display the customer details on the form
        txtName.Text = customer.Name
        txtFirstName.Text = customer.FirstName
        txtLastName.Text = customer.LastName
        txtEmail.Text = customer.Email
    End Sub

    Private Sub btnListCustomer_Click(sender As Object, e As EventArgs) Handles btnListCustomer.Click
        'Create a new customer
        'Dim objCustomer As Customer
        'objCustomer.FirstName = "Michael"
        'objCustomer.LastName = "Dell"
        'objCustomer.Email = "mdell@somecompany.com"

        'Display the customer
        'DisplayCustomer(objCustomer)

        'Create some customers
        CreateCustomer("Darrel", "Hilton", "dhilton@somecompany.com")
        CreateCustomer("Frank", "Peoples", "fpeoples@somecompany.com")
        CreateCustomer("Bill", "Scott", "bscott@somecompany.com")
        CreateCustomer("Michael", "Dell", "mdell@somecompany.com")

        'btnListCustomer.Enabled = False

    End Sub

    Private Sub btnDeleteCustomer_Click(sender As Object, e As EventArgs) Handles btnDeleteCustomer.Click
        'If no customer is selected in the ListBox then.
        If lstCustomers.SelectedIndex = -1 Then
            'Display a message
            MessageBox.Show("You must select a customer to delete.", "Structure Demo")

            'Exit this method
            Exit Sub
        End If

        'Prompt the user to delete the selected customer
        If MessageBox.Show("Are you sure you want to delete " & SelectedCustomer.Name & "?", "Structure Demo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            'Get the customer to be deleted
            Dim objCustomerToDelete As Customer = SelectedCustomer

            'Remove the customer from the ArrayList
            objCustomers.Remove(objCustomerToDelete)

            'Remove the customer from the ListBox
            lstCustomers.Items.Remove(objCustomerToDelete)
        End If
    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCustomers.SelectedIndexChanged

        'Display the customer details
        DisplayCustomer(SelectedCustomer)
    End Sub

    Private Sub btnLookup_Click(sender As Object, e As EventArgs) Handles btnLookup.Click
        'Declare a customer object and set it to the customer
        'with the email address to be found
        Dim objFoundCustomer As Customer = objCustomers(txtEmail.Text)
        If Not IsNothing(objFoundCustomer.Email) Then
            'Display the customers name
            MessageBox.Show("The customers name is: " &
            objFoundCustomer.Name, "Structure Demo")
        Else
            'Display an error
            MessageBox.Show("There is no customer with the e-mail" &
            " address " & txtEmail.Text & ".", "Structure Demo")
        End If
    End Sub

    Private Sub btnAddCustomer_Click(sender As Object, e As EventArgs) Handles btnAddCustomer.Click
        Dim objFoundCustomer As Customer = objCustomers(txtEmail.Text)
        If Not IsNothing(objFoundCustomer.Email) Then
            'Display the customers name
            MessageBox.Show("There is already a customer (" &
            objFoundCustomer.Name & ") with this email address", "Structure Demo")
        Else
            CreateCustomer(txtFirstName.Text, txtLastName.Text, txtEmail.Text)
            txtFirstName.Text = ""
            txtLastName.Text = ""
            txtEmail.Text = ""
        End If
    End Sub

    Private Sub txtLastName_TextChanged(sender As Object, e As EventArgs) Handles txtLastName.TextChanged
        NameChanged(txtFirstName.Text, txtLastName.Text)
    End Sub

    Private Sub txtFirstName_TextChanged(sender As Object, e As EventArgs) Handles txtFirstName.TextChanged
        NameChanged(txtFirstName.Text, txtLastName.Text)
    End Sub

    Private Sub NameChanged(FirstName As String, LastName As String)
        txtName.Text = FirstName & Space(1) & LastName
    End Sub

    Private Sub btnDeleteAll_Click(sender As Object, e As EventArgs) Handles btnDeleteAll.Click
        'Clear the lists
        objCustomers.Clear()
        lstCustomers.Items.Clear()
    End Sub
End Class
