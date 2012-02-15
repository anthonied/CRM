using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace SolExcelCrmSync.Classes
{
    public class Invoice
    {
        string _DebtorAccountNumber;
        [CsvColumn(Name = "DebtorAccountNumber", FieldIndex=0)]        
        public string DebtorAccountNumber
        {
            
            get { return _DebtorAccountNumber; }
            set
            {
                int maxlength = 5;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }

                _DebtorAccountNumber = value;
            }
        }

        string _OrderNumber;
        [CsvColumn(Name= "OrderNumber", FieldIndex=1)]
        public string OrderNumber
        {
            get { return _OrderNumber; }
            set { _OrderNumber = value; }
        }

        string _InvoiceNumber;
        [CsvColumn(Name = "InvoiceNumber", FieldIndex = 2)]
        public string InvoiceNumber
        {             
            get { return _InvoiceNumber; }
            set
            {
                int maxlength = 8;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _InvoiceNumber = value;
            }
        }

        DateTime _InvoiceDate;
        [CsvColumn(Name="InvoiceDate",  FieldIndex=3, OutputFormat="ddMMyy")]
        public DateTime InvoiceDate
        {
            get { return _InvoiceDate; }
            set { _InvoiceDate = value; }
        }

        string _CustomerReference;
        [CsvColumn(Name = "CustomerReference", FieldIndex = 4)]
        public string CustomerReference
        {
            get { return _CustomerReference; }
            set
            {
                int maxlength = 10;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _CustomerReference = value;
            }
        }

        string _JobNumber;
       
        public string JobNumber        {
            get { return _JobNumber; }
            set { _JobNumber = value; }
        }

        string _Consultant;
        [CsvColumn(Name = "Consultant", FieldIndex = 5)]
        public string Consultant
        {
            get { return _Consultant; }
            set { _Consultant = value; }
        }

        string _ShippingMethod;
        [CsvColumn(Name = "ShippingMethod", FieldIndex = 6)]
        public string ShippingMethod
        {
            get { return _ShippingMethod; }
            set
            {
                int maxlength = 30;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _ShippingMethod = value;
            }
        }

        string _DeliveryLine1;
        [CsvColumn(Name = "DeliveryLine1", FieldIndex = 7)]
        public string DeliveryLine1
        {
            get { return _DeliveryLine1; }
            set
            {
                int maxlength = 30;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _DeliveryLine1 = value; 
            }
        }

        string _DeliveryLine2;
        [CsvColumn(Name = "DeliveryLine2", FieldIndex = 8)]
        public string DeliveryLine2
        {
            get { return _DeliveryLine2; }
            set { _DeliveryLine2 = value; }
        }

        string _DeliveryLine3;
        [CsvColumn(Name = "DeliveryLine3", FieldIndex = 9)]
        public string DeliveryLine3
        {
            get { return _DeliveryLine3; }
            set
            {
                int maxlength = 30;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _DeliveryLine3 = value;
            }
        }

        string _DeliveryLine4;
        [CsvColumn(Name = "DeliveryLine4", FieldIndex =10)]
        public string DeliveryLine4
        {
            get { return _DeliveryLine4; }
            set
            {
                int maxlength = 30;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _DeliveryLine4 = value;
            }
        }

        string _SpecialInstructions1;
        [CsvColumn(Name = "SpecialInstructions1", FieldIndex = 11)]
        public string SpecialInstructions1
        {
            get { return _SpecialInstructions1; }
            set { _SpecialInstructions1 = value; }
        }

        string _SpecialInstructions2;
        [CsvColumn(Name = "SpecialInstructions2", FieldIndex = 12)]
        public string SpecialInstructions2
        {
            get { return _SpecialInstructions2; }
            set
            {
                int maxlength = 30;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _SpecialInstructions2 = value;
            }
        }

        string _Product_String;
        [CsvColumn(Name = "ProductString", FieldIndex = 13)]
        public string Product_String
        {
            get { return _Product_String; }
            set { _Product_String = value; }
        }


        List<Invoice_detail> _Invoice_Product;

        public List<Invoice_detail> Invoice_Product
        {
            get { return _Invoice_Product; }
            set { _Invoice_Product = value; }
        }
       
       /* string _StatusCode;        
        public string StatusCode
        {
            get { return _StatusCode; }
            set { _StatusCode = value; }
        }*/
    }

    public class Invoice_detail
    {
        private string _ProductCode;
        public string ProductCode
        {
            get { return _ProductCode; }
            set { _ProductCode = value; }
        }

        private string _ExtendedDescription;
        public string ExtendedDescription
        {
            get { return _ExtendedDescription; }
            set { _ExtendedDescription = value; }
        }

        private string _TextNarrativeLine1;
        public string TextNarrativeLine1
        {
            get { return _TextNarrativeLine1; }
            set 
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine1 = value; 
            }
        }

        private string _TextNarrativeLine2;
        public string TextNarrativeLine2
        {
            get { return _TextNarrativeLine2; }
            set
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine2 = value; 
            }
        }

        private string _TextNarrativeLine3;
        public string TextNarrativeLine3
        {
            get { return _TextNarrativeLine3; }
            set
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine3 = value; 
            }
        }

        private string _TextNarrativeLine4;
        public string TextNarrativeLine4
        {
            get { return _TextNarrativeLine4; }
            set 
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine4 = value;
            }
        }

        private string _TextNarrativeLine5;
        public string TextNarrativeLine5
        {
            get { return _TextNarrativeLine5; }
            set 
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine5 = value;
            }
        }

        private string _TextNarrativeLine6;
        public string TextNarrativeLine6
        {
            get { return _TextNarrativeLine6; }
            set
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine6 = value; 
            }
        }

        private string _TextNarrativeLine7;
        public string TextNarrativeLine7
        {
            get { return _TextNarrativeLine7; }
            set
            {
                int maxlength = 60;
                if (value.Length > maxlength)
                {
                    value = value.Substring(0, maxlength);
                }
                _TextNarrativeLine7 = value; 
            }
        }

        private string _JobNumber_EarnBy;
        public string JobNumber_EarnBy
        {
            get { return _JobNumber_EarnBy; }
            set { _JobNumber_EarnBy = value; }
        }

        private string _SerialNumber_SoldBy;
        public string SerialNumber_SoldBy
        {
            get { return _SerialNumber_SoldBy; }
            set { _SerialNumber_SoldBy = value; }
        }

        private string _NomCostCenter_GLCC;
        public string NomCostCenter_GLCC
        {
            get { return _NomCostCenter_GLCC; }
            set { _NomCostCenter_GLCC = value; }
        }

        private string _GLAC;
        public string GLAC
        {
            get { return _GLAC; }
            set { _GLAC = value; }
        }

        private decimal _Quantity;
        public decimal Quantity
        {
            get { return _Quantity; }
            set { _Quantity = value; }
        }

        private decimal _UnitPrice;
        public decimal UnitPrice
        {
            get { return _UnitPrice; }
            set { _UnitPrice = value; }
        }
    }
}
