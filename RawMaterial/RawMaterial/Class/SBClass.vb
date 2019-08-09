Imports System.Text
Imports DJLib.Dbtools
Public Class SBClass

    Public Shared Function GenSBCurrenciesLine(ByVal Date1 As Date, ByVal Date2 As Date, Optional ByVal point As Integer = 1) As StringBuilder
        Dim stringbuilder1 As New StringBuilder
        stringbuilder1.Append("select crcydate::text as ""TX Date"",  " & point & "as ""Amount"" from rmcrcytx ")
        stringbuilder1.Append("where crcydate >= ")
        stringbuilder1.Append(DateFormatyyyyMMdd(Date1))
        stringbuilder1.Append("and crcydate <= ")
        stringbuilder1.Append(DateFormatyyyyMMdd(Date2))
        stringbuilder1.Append("group by crcydate order by ""TX Date""")
        Return stringbuilder1
    End Function

    Public Shared Function GenSBCurrencies(ByVal Crcy1 As String, ByVal Crcy2 As String, ByVal Date1 As Date, ByVal Date2 As Date, Optional ByVal Basepoint As Decimal = 1) As StringBuilder
        Dim stringbuilder1 As New StringBuilder
        stringbuilder1.Append("select rm.crcydate::text as ""TX Date"",(" & IIf(Crcy1 = "EUR", 1, "rm.crcyamount") & " / " & IIf(Crcy2 = "EUR", 1, "rm2.crcyamount") & " * " & Basepoint & ") as ""Amount"" from rmcrcytx rm")
        stringbuilder1.Append(" left join rmcrcytx rm2 on rm2.crcydate = rm.crcydate and rm2.crcycode = '" & Crcy2 & "'")
        stringbuilder1.Append(" where rm.crcydate>=") '2010-01-01'  
        stringbuilder1.Append(DateFormatyyyyMMdd(Date1))
        stringbuilder1.Append(" and rm.crcydate <=")  '2010-10-01' 
        stringbuilder1.Append(DateFormatyyyyMMdd(Date2))
        stringbuilder1.Append(" and rm.crcycode = '" & IIf(Crcy1 = "EUR", Crcy2, Crcy1) & "'")
        stringbuilder1.Append(" and not " & IIf(Crcy1 = "EUR", 1, "rm.crcyamount") & " / " & IIf(Crcy2 = "EUR", 1, "rm2.crcyamount") & " isnull")
        stringbuilder1.Append(" order by rm.crcydate")
        Return stringbuilder1
    End Function

    Public Shared Function GenSBRM(ByVal periodType As Integer, ByVal CrcyTo As String, ByVal rmId As String, ByVal Date1 As Date, ByVal Date2 As Date, Optional ByVal BasePoint As Decimal = 0) As StringBuilder
        Dim stringbuilder1 As New StringBuilder

        'stringbuilder1.Append("select txdate::text as ""Tx Date"",txamount as ""Original Amount"",txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric  / txamount::numeric as conversion ,tx.crcycode  as ""Original Crcy"", txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric  as ""Tx Amount"" ,txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric * " & BasePoint & "::numeric  as ""Tx Amount (Base 100)"", '" & CrcyTo & "' as ""Crcy""  from rmmaterialrawtx tx ")
        stringbuilder1.Append("select txdate::text as ""Tx Date"",txamount as ""Original Amount"",")
        stringbuilder1.Append("case when '" & CrcyTo & "' = tx.crcycode then 1 else ")
        stringbuilder1.Append("txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric  / txamount::numeric ")
        stringbuilder1.Append(" end as conversion ,")
        stringbuilder1.Append("tx.crcycode  as ""Original Crcy"",")
        stringbuilder1.Append("case when '" & CrcyTo & "' = tx.crcycode then 1 * txamount else ")
        stringbuilder1.Append(" txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric")
        stringbuilder1.Append(" end as ""Tx Amount"" ,")
        stringbuilder1.Append("case when '" & CrcyTo & "' = tx.crcycode then 1 * txamount * " & BasePoint & " else ")
        stringbuilder1.Append("txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric * " & BasePoint & "::numeric")
        stringbuilder1.Append(" end as ""Tx Amount (Base 100)"",")
        stringbuilder1.Append(" '" & CrcyTo & "' as ""Crcy""  from rmmaterialrawtx tx ")
        Select Case periodType
            Case 1
                stringbuilder1.Append(" left join rmcrcytx c1 on c1.crcydate = tx.txdate and c1.crcycode= tx.crcycode")
                stringbuilder1.Append(" left join rmcrcytx c2 on c2.crcydate = tx.txdate and c2.crcycode= '")
            Case 7
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.myweek = date_part('week',txdate) and c1.crcycode = tx.crcycode ")
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.myweek = date_part('week',txdate) and c2.crcycode = '")
            Case 31
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.mymonth = date_part('month',txdate) and c1.crcycode = tx.crcycode ")
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.mymonth = date_part('month',txdate) and c2.crcycode = '")

        End Select

        stringbuilder1.Append(CrcyTo)
        stringbuilder1.Append("'")

        stringbuilder1.Append(" where keyid = '")
        stringbuilder1.Append(rmId)
        stringbuilder1.Append("' and txdate >= ")
        stringbuilder1.Append(DateFormatyyyyMMdd(Date1))
        stringbuilder1.Append(" and txdate <= ")
        stringbuilder1.Append(DateFormatyyyyMMdd(Date2))
        stringbuilder1.Append(" order by txdate")
        Return stringbuilder1
    End Function

    'Public Shared Function GenSBRMSingle(ByVal periodType As Integer, ByVal CrcyTo As String, ByVal rmId As String, ByVal Date1 As Date, ByVal Date2 As Date, Optional ByVal BasePoint As Decimal = 1) As StringBuilder
    Public Shared Function GenSBRMSingle(ByVal periodType As Integer, ByVal CrcyTo As String, Optional ByVal BasePoint As Decimal = 1) As StringBuilder
        Dim stringbuilder1 As New StringBuilder
        'stringbuilder1.Append("select txdate::date as ""Tx Date"",txamount as ""Original Amount"",txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric  / txamount::numeric as conversion,tx.crcycode  as ""Original Crcy"", txamount / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric * " & BasePoint & "  as ""Tx Amount (" & CrcyTo & ")""" & " , '" & CrcyTo & "' as ""Crcy"" ,hd.rawmaterialname || ' (' || hd.source || ')'  || ' (' || tx.crcycode || ')' as rawmaterialname from rmmaterialrawtx tx ")
        stringbuilder1.Append("select txdate::date as ""Tx Date"",txamount as ""Original Amount"",")
        stringbuilder1.Append("case when '" & CrcyTo & "' = tx.crcycode then 1 else ")
        stringbuilder1.Append("txamount::numeric / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric  * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric  / txamount::numeric")                             
        stringbuilder1.Append(" end as conversion ,")
        stringbuilder1.Append("tx.crcycode  as ""Original Crcy"",")
        stringbuilder1.Append("case when '" & CrcyTo & "' = tx.crcycode then 1 * txamount * " & BasePoint & " else ")
        stringbuilder1.Append(" txamount / validatecrcyamount(c1.crcyamount,tx.txdate,tx.crcycode)::numeric * validatecrcyamount(c2.crcyamount,tx.txdate,'" & CrcyTo & "')::numeric * " & BasePoint)
        stringbuilder1.Append(" end as ""Tx Amount (" & CrcyTo & ")""" & " ,")
        stringbuilder1.Append(" '" & CrcyTo & "' as ""Crcy"" ,hd.rawmaterialname || ' (' || hd.source || ' - ' || hd.unit || ')'  || ' (' || tx.crcycode || ')' as rawmaterialname from rmmaterialrawtx tx ")
        Select Case periodType
            Case 1
                stringbuilder1.Append(" left join rmcrcytx c1 on c1.crcydate = tx.txdate and c1.crcycode= tx.crcycode")
                stringbuilder1.Append(" left join rmcrcytx c2 on c2.crcydate = tx.txdate and c2.crcycode= '")
            Case 7
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.myweek = date_part('week',txdate) and c1.crcycode = tx.crcycode ")
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.myweek = date_part('week',txdate) and c2.crcycode = '")
            Case 31
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.mymonth = date_part('month',txdate) and c1.crcycode = tx.crcycode ")
                stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.mymonth = date_part('month',txdate) and c2.crcycode = '")
        End Select

        stringbuilder1.Append(CrcyTo)
        stringbuilder1.Append("'")
        stringbuilder1.Append(" left join rmmaterialraw hd on hd.keyid = tx.keyid ")

        'stringbuilder1.Append(" where ")
        'stringbuilder1.Append(rmId)
        'If Date1.Year > 1 Then
        '    stringbuilder1.Append("' and txdate >= ")
        '    stringbuilder1.Append(DateFormatyyyyMMdd(Date1))
        'End If
        'If Date2.Year > 1 Then
        '    stringbuilder1.Append(" and txdate <= ")
        '    stringbuilder1.Append(DateFormatyyyyMMdd(Date2))
        'End If
        'stringbuilder1.Append(" order by txdate")

        Return stringbuilder1
    End Function
End Class
