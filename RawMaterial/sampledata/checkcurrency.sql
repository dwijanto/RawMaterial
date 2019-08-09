select txdate::text as "Tx Date",txamount as "Original Amount",
txamount / ( case when tx.crcycode = 'EUR' then 1 else 
case when
	c1.crcyamount > 0 then c1.crcyamount
	else getlastcurrency(tx.txdate,tx.crcycode)
  end  
end )  * 1  / txamount as conversion ,
tx.crcycode  as "Original Crcy", 
txamount / ( case when tx.crcycode = 'EUR' then 1 else 
  case when
	c1.crcyamount > 0 then c1.crcyamount
	else getlastcurrency(tx.txdate,tx.crcycode)
  end 
end)  * 1  as "Tx Amount" ,
txamount / ( case when tx.crcycode = 'EUR' then 1 else 
case when
	c1.crcyamount > 0 then c1.crcyamount
	else getlastcurrency(tx.txdate,tx.crcycode)
  end 

end )  * 1 * 1  as "Tx Amount (Base 100)"
, 'EUR' as "Crcy"  
from rmmaterialrawtx tx  
left join rmcrcytx c1 on c1.crcydate = tx.txdate and c1.crcycode= tx.crcycode 
left join rmcrcytx c2 on c2.crcydate = tx.txdate and c2.crcycode= 'EUR' 
where keyid = 'NF001' and txdate >= '2006-10-8' and txdate <= '2010-10-8' order by txdate