Public Class Form1

    'Form level members
    Private objCustomers As New CustomerCollection

    Public Sub CreateCustomer(ByVal firstName As String, _
    ByVal lastName As String, ByVal email As String)

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

    Private Sub btnListCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnListCustomer.Click

        'Clear the lists
        objCustomers.Clear()
        lstCustomers.Items.Clear()

        'Create some customers
        CreateCustomer("Darrel", "Hilton", "dhilton@somecompany.com")
        CreateCustomer("Frank", "Peoples", "fpeoples@somecompany.com")
        CreateCustomer("Bill", "Scott", "bscott@somecompany.com")

    End Sub

    Public ReadOnly Property SelectedCustomer() As Customer
        Get
            If lstCustomers.SelectedIndex <> -1 Then
                'Return the selected customer
                Return objCustomers(lstCustomers.SelectedIndex)
            End If
        End Get
    End Property

    Private Sub btnDeleteCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCustomer.Click

        'If no customer is selected in the ListBox then.
        If lstCustomers.SelectedIndex = -1 Then

            'Display a message
            MessageBox.Show("You must select a customer to delete.", _
                "Structure Demo")

            'Exit this method
            Exit Sub
        End If

        'Prompt the user to delete the selected customer
        If MessageBox.Show("Are you sure you want to delete " & _
            SelectedCustomer.Name & "?", "Structure Demo", _
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) = _
            DialogResult.Yes Then

            'Get the customer to be deleted
            Dim objCustomerToDelete As Customer = SelectedCustomer

            'Remove the customer from the ArrayList
            objCustomers.Remove(objCustomerToDelete)

            'Remove the customer from the ListBox
            lstCustomers.Items.Remove(objCustomerToDelete)
        End If

    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCustomers.SelectedIndexChanged

        'Display the customer details
        DisplayCustomer(SelectedCustomer)

    End Sub

    Private Sub btnLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLookup.Click

        'Declare a customer object and set it to the customer
        'with the email address to be found
        Dim objFoundCustomer As Customer = objCustomers(txtEmail.Text)

        If Not IsNothing(objFoundCustomer.Email) Then
            'Display the customers name
            MessageBox.Show("The customers name is: " & _
                objFoundCustomer.Name, "Structure Demo")
        Else
            'Display an error
            MessageBox.Show("There is no customer with the e-mail" & _
                " address " & txtEmail.Text & ".", "Structure Demo")
        End If

    End Sub
End Class
