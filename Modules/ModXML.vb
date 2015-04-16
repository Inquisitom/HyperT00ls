Imports System.IO
Imports System.Xml
Imports System.Text.RegularExpressions

Module ModXML
    Public Structure RomEntry
        Dim RomName As String
        Dim Description As String
        Dim Enabled As Boolean
        Dim Image As String
        Dim Index As String
        Dim CRC As String
        Dim Manufacturer As String
        Dim Year As String
        Dim Genre As String
        Dim Rating As String
        Dim CloneOf As String
    End Structure

    Public Sub FillRomsDataGrid_From_SystemName(ByVal SystemName As String, ByRef vDataGrid As DataGridView)
        Dim XMLSystem As String = gCONFIGFILE
    End Sub

    Function RomsInfos(ByVal XMLFullPath As String) As RomEntry
        Dim vROM As New RomEntry
        Try
            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode
            'Create the XML Document
            m_xmld = New XmlDocument()
            'Load the Xml file
            m_xmld.Load(XMLFullPath)
            'Get the list of name nodes 
            m_nodelist = m_xmld.SelectNodes("/game")
            'Loop through the nodes
            For Each m_node In m_nodelist
                'Get the Gender Attribute Value
                vROM.Description = m_node.Attributes.GetNamedItem("Description").Value
                'Get the firstName Element Value
                Dim firstNameValue = m_node.ChildNodes.Item(0).InnerText
                'Get the lastName Element Value
                Dim lastNameValue = m_node.ChildNodes.Item(1).InnerText
            Next
        Catch errorVariable As Exception
            'Error trapping
        End Try
        Return vROM
    End Function

    Public Function EncodeForXml(ByVal data As String) As String
        Static badAmpersand As New Regex("&(?![a-zA-Z]{2,6};|#[0-9]{2,4};)")

        data = badAmpersand.Replace(data, "&amp;")

        Return data.Replace("<", "&lt;").Replace("""", "&quot;").Replace(">", "gt;")
    End Function
End Module
