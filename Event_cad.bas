Attribute VB_Name = "Event_cad"
Function get_point() As Variant
code_cad.GetObject_Acad
get_point = Acad.ActiveDocument.Utility.GetPoint(, "chon 1 diem")
End Function
  ?        8?O  ?        ;?O  ?s  (   :?O 	 ?        =?O 
 ?        <?O  ?        ??O  ?        >?O  ?        1?O  ?        0?O  ?t  (   3?O  ?        2?	O  ?        5?
O  ?u  (   4?O  ?        7?O  ?        6?O  ?        	?O  ?        ?O  ?        ? O  ?        
?O  ?        ?O  ?                