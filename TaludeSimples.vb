Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime 'commandMethod()
Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application

Public Class TaludeSimples
    <CommandMethod("TALUDESIMPLES")>
    Public Sub TaludeSimples()
        'pega o documento
        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        'a database 
        Dim db = doc.Database
        Dim ed = doc.Editor

        Using tr = db.TransactionManager.StartTransaction
            'blocktable e model
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim model As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            'pega o polyline
            Dim crista As Polyline = tr.GetObject(ObjectSelIdClick("Selecione a crista", GetType(Polyline)), OpenMode.ForRead)
            Dim pe As Polyline = tr.GetObject(ObjectSelIdClick("Selecione o pé", GetType(Polyline)), OpenMode.ForRead)
            'pergunta a frequencia
            Dim espacamento As Double = ed.GetDouble("Qual o espaçamento?").Value

            'itera de 0 ate o final conforme a frenquencia
            For i = 0 To crista.Length Step espacamento

                'pega o ponto na crista na distancia i
                Dim ponto As Point3d = crista.GetPointAtDist(i)
                'vetor tangente a crista
                Dim tangente As Vector3d = crista.GetFirstDerivative(ponto)
                'vetor perpendicular
                Dim perpendicular As Vector3d = tangente.GetPerpendicularVector

                'ponto no pe saindo na perpendicular do ponto da crista
                Dim pontoPe As Point3d = pe.GetClosestPointTo(ponto, perpendicular, False)

                'vetor e ponto intermediario
                Dim vetorInter As Vector3d = (pontoPe - ponto)
                Dim pontoInter As Point3d = ponto + vetorInter * 0.5

                'cria uma polyline nos pontos
                Dim poly As New Polyline
                poly.AddVertexAt(poly.NumberOfVertices, New Point2d(ponto.X, ponto.Y), 0, 0, 0)

                'se for impar usa o ponto intermediario
                If (i Mod 2) Then
                    poly.AddVertexAt(poly.NumberOfVertices, New Point2d(pontoInter.X, pontoInter.Y), 0, 0, 0)
                Else
                    poly.AddVertexAt(poly.NumberOfVertices, New Point2d(pontoPe.X, pontoPe.Y), 0, 0, 0)
                End If
                'adiciona a poyline ao model
                model.AppendEntity(poly)
                tr.AddNewlyCreatedDBObject(poly, True)

            Next
            'confirma a alteracao
            tr.Commit()

        End Using

    End Sub
    ''' <summary>
    ''' Função que retorna um unico objeto selecionado, mostra mensagem específica ao usuario para selecao com filtro para seleção.
    ''' </summary>
    Public Function ObjectSelIdClick(msgSelecao As String, tipoClass As Type) As ObjectId
        Dim doc = acApp.DocumentManager.MdiActiveDocument
        'pede objeto
        Dim peo As New PromptEntityOptions(vbLf & msgSelecao)
        peo.SetRejectMessage(msgSelecao & " !!!")
        peo.AddAllowedClass(tipoClass, True)
        Dim per As PromptEntityResult = doc.Editor.GetEntity(peo)
        If per.Status <> PromptStatus.OK Then Return Nothing
        Return per.ObjectId
    End Function
End Class
