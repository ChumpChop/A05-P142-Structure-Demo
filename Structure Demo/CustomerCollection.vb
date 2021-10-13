﻿Public Class CustomerCollection
    Inherits CollectionBase

    'Private member
    Private objEmailHashtable As New Hashtable

    'EmailHashtable property to return the Email Hashtable
    Public ReadOnly Property EmailHashtable() As Hashtable
        Get
            Return objEmailHashtable
        End Get
    End Property

    'Item property to read or update a customer at a given position
    'in the list
    Default Public Property Item(ByVal index As Integer) As Customer
        Get
            Return CType(Me.List.Item(index), Customer)
        End Get
        Set(ByVal value As Customer)
            Me.List.Item(index) = value
        End Set
    End Property

    'Add a customer to the collection
    Public Sub Add(ByVal newCustomer As Customer)
        Me.List.Add(newCustomer)

        'Add the email address to the Hashtable
        EmailHashtable.Add(newCustomer.Email.ToLower, newCustomer)
    End Sub


    'Remove a customer from the collection
    Public Sub Remove(ByVal oldCustomer As Customer)
        Me.List.Remove(oldCustomer)

        'Remove customer from the Hashtable
        EmailHashtable.Remove(oldCustomer.Email.ToLower)
    End Sub


    'Provide a new implementation of the RemoveAt method
    Public Shadows Sub RemoveAt(ByVal index As Integer)
        Remove(Item(index))
    End Sub

    'Overload Item property to find a customer by email address
    Default Public ReadOnly Property Item(ByVal email As String) As Customer
        Get
            Return CType(EmailHashtable.Item(email.ToLower), Customer)
        End Get
    End Property


    'Provide a new implementation of the Clear method
    Public Shadows Sub Clear()
        'Clear the CollectionBase
        MyBase.Clear()
        'Clear your hashtable
        EmailHashtable.Clear()
    End Sub

End Class
