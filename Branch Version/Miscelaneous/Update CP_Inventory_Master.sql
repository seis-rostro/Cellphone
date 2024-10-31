SELECT a.sstockidx, a.nbegqtyxx, a.nqtyonhnd 
From CP_Inventory_Master a  
   LEFT JOIN CP_Inventory b  
on a.sstockidx = b.sstockidx
Where a.sbranchcd = '09'  
   AND a.nqtyonhnd <> 0  
   AND b.sCategIDx = '01001'  
