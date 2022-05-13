Namespace Enums.ImageMSO

    Public NotInheritable Class Mail
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property MailMergeAddressBlockInsert As Mail = New Mail(1, NameOf(MailMergeAddressBlockInsert))
        Public Shared ReadOnly Property MailMergeAutoCheckForErrors As Mail = New Mail(2, NameOf(MailMergeAutoCheckForErrors))
        Public Shared ReadOnly Property MailMergeClearMergeType As Mail = New Mail(3, NameOf(MailMergeClearMergeType))
        Public Shared ReadOnly Property MailMergeCreateList As Mail = New Mail(4, NameOf(MailMergeCreateList))
        Public Shared ReadOnly Property MailMergeDocument As Mail = New Mail(5, NameOf(MailMergeDocument))
        Public Shared ReadOnly Property MailMergeFindRecipient As Mail = New Mail(6, NameOf(MailMergeFindRecipient))
        Public Shared ReadOnly Property MailMergeFinishAndMergeMenu As Mail = New Mail(7, NameOf(MailMergeFinishAndMergeMenu))
        Public Shared ReadOnly Property MailMergeGoToFirstRecord As Mail = New Mail(8, NameOf(MailMergeGoToFirstRecord))
        Public Shared ReadOnly Property MailMergeGoToNextRecord As Mail = New Mail(9, NameOf(MailMergeGoToNextRecord))
        Public Shared ReadOnly Property MailMergeGoToPreviousRecord As Mail = New Mail(10, NameOf(MailMergeGoToPreviousRecord))
        Public Shared ReadOnly Property MailMergeGotToLastRecord As Mail = New Mail(11, NameOf(MailMergeGotToLastRecord))
        Public Shared ReadOnly Property MailMergeGreetingLineInsert As Mail = New Mail(12, NameOf(MailMergeGreetingLineInsert))
        Public Shared ReadOnly Property MailMergeHelper As Mail = New Mail(13, NameOf(MailMergeHelper))
        Public Shared ReadOnly Property MailMergeHighlightMergeFields As Mail = New Mail(14, NameOf(MailMergeHighlightMergeFields))
        Public Shared ReadOnly Property MailMergeMatchFields As Mail = New Mail(15, NameOf(MailMergeMatchFields))
        Public Shared ReadOnly Property MailMergeMergeFieldInsert As Mail = New Mail(16, NameOf(MailMergeMergeFieldInsert))
        Public Shared ReadOnly Property MailMergeMergeFieldInsertMenu As Mail = New Mail(17, NameOf(MailMergeMergeFieldInsertMenu))
        Public Shared ReadOnly Property MailMergeMergeToDocument As Mail = New Mail(18, NameOf(MailMergeMergeToDocument))
        Public Shared ReadOnly Property MailMergeMergeToEMail As Mail = New Mail(19, NameOf(MailMergeMergeToEMail))
        Public Shared ReadOnly Property MailMergeMergeToFax As Mail = New Mail(20, NameOf(MailMergeMergeToFax))
        Public Shared ReadOnly Property MailMergeMergeToPrinter As Mail = New Mail(21, NameOf(MailMergeMergeToPrinter))
        Public Shared ReadOnly Property MailMergeRecepientsUseExistingList As Mail = New Mail(22, NameOf(MailMergeRecepientsUseExistingList))
        Public Shared ReadOnly Property MailMergeRecepientsUseOutlookContacts As Mail = New Mail(23, NameOf(MailMergeRecepientsUseOutlookContacts))
        Public Shared ReadOnly Property MailMergeRecipientsEditList As Mail = New Mail(24, NameOf(MailMergeRecipientsEditList))
        Public Shared ReadOnly Property MailMergeResultsPreview As Mail = New Mail(25, NameOf(MailMergeResultsPreview))
        Public Shared ReadOnly Property MailMergeRules As Mail = New Mail(26, NameOf(MailMergeRules))
        Public Shared ReadOnly Property MailMergeSelectRecipients As Mail = New Mail(27, NameOf(MailMergeSelectRecipients))
        Public Shared ReadOnly Property MailMergeSetDocumentType As Mail = New Mail(28, NameOf(MailMergeSetDocumentType))
        Public Shared ReadOnly Property MailMergeStartDirectory As Mail = New Mail(29, NameOf(MailMergeStartDirectory))
        Public Shared ReadOnly Property MailMergeStartEmail As Mail = New Mail(30, NameOf(MailMergeStartEmail))
        Public Shared ReadOnly Property MailMergeStartEnvelopes As Mail = New Mail(31, NameOf(MailMergeStartEnvelopes))
        Public Shared ReadOnly Property MailMergeStartLabels As Mail = New Mail(32, NameOf(MailMergeStartLabels))
        Public Shared ReadOnly Property MailMergeStartLetters As Mail = New Mail(33, NameOf(MailMergeStartLetters))
        Public Shared ReadOnly Property MailMergeStartMailMergeMenu As Mail = New Mail(34, NameOf(MailMergeStartMailMergeMenu))
        Public Shared ReadOnly Property MailMergeUpdateLabels As Mail = New Mail(35, NameOf(MailMergeUpdateLabels))
        Public Shared ReadOnly Property MailMergeWizard As Mail = New Mail(36, NameOf(MailMergeWizard))
        Public Shared ReadOnly Property MailMove As Mail = New Mail(37, NameOf(MailMove))
        Public Shared ReadOnly Property MailSelectNames As Mail = New Mail(38, NameOf(MailSelectNames))


        Public Overrides Function Clone() As Object
            Return New Mail(value, name)
        End Function

    End Class

End Namespace