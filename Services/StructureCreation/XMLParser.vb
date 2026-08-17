Namespace Abovo
    Public Class XMLParser

        Private Sub ReadXml()
            Dim xmlAll = <?xml version="1.0" encoding="windows-1252"?>
                         <MatML_Doc>
                             <Material>
                                 <BulkDetails>
                                     <Name>23133385</Name>
                                     <Class>
                                         <Name>1 - Carbon Steel</Name>
                                     </Class>
                                     <Source source=""/>
                                     <PropertyData property="Material Type">
                                         <Data format="string">IsotropicMaterial</Data>
                                     </PropertyData>
                                     <PropertyData Property="Mass Density (RHO)_1">
                                         <Data format="exponential">7.87e-6</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Organization">
                                         <Data format="string">SAE</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Name">
                                         <Data format="string">J1199</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Grade">
                                         <Data format="string">CLASS 9.8</Data>
                                     </PropertyData>
                                 </BulkDetails>
                             </Material>
                             <Material>
                                 <BulkDetails>
                                     <Name>23133419</Name>
                                     <Class>
                                         <Name>1 - Carbon Steel</Name>
                                     </Class>
                                     <Source source=""/>
                                     <PropertyData property="Material Type">
                                         <Data format="string">IsotropicMaterial</Data>
                                     </PropertyData>
                                     <PropertyData Property="Mass Density (RHO)_1">
                                         <Data format="exponential">7.87e-6</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Organization">
                                         <Data format="string">EN</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Name">
                                         <Data format="string">10130</Data>
                                     </PropertyData>
                                     <PropertyData Property="Spec Grade">
                                         <Data format="string">DC05</Data>
                                     </PropertyData>
                                 </BulkDetails>
                             </Material>
                         </MatML_Doc>

            For Each xel In xmlAll.Elements
                'Dim material As New MyMaterial
                'material.Name = xel.Element("BulkDetails").Element("Name").Value
                'material.Classe = xel.Element("BulkDetails").Element("Classe").Element("Name").Value
                'material.Org = xel.Element("BulkDetails").Elements.Where(Function(d) d.Name = "PropertyData" And d.Attribute("property").Value = "Material Type").Value
                ''Ect
            Next

        End Sub

    End Class

End Namespace
