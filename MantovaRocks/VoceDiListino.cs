using System;
using System.Linq;
using Xbim.Ifc4.Interfaces;

namespace MantovaRocks
{
	internal class VoceDiListino
	{
		public Double Costo;
		public TipoDiCosto LogicaDiCalcolo;

		internal double Calcola(IIfcBuildingElement objectToCompute, out string errore)
		{
			errore = "";
			if (LogicaDiCalcolo == TipoDiCosto.ACorpo)
				return Costo;

			// cerchiamo la quantita' associata alla logica di calcolo
			// 
			var quantities = objectToCompute.IsDefinedBy.FirstOrDefault(x => x.RelatingPropertyDefinition is IIfcElementQuantity);
			if (quantities == null)
			{
				errore = "quantities not Found";
				return 0;
			}

			var t = quantities.RelatingPropertyDefinition as IIfcElementQuantity;
			if (t == null)
			{
				errore = "invalid quantities";
				return 0;
			}
			
			var trovata = t.Quantities.FirstOrDefault(x => x.Name.ToString() == MatchName());
			if (trovata == null)
			{
				var names = string.Join(",", t.Quantities.Select(x => x.Name.ToString()));
				errore = $"quantity {MatchName()} not found; quantities available are: {names}.";
				return 0;
			}
			if (trovata is IIfcQuantityArea qa)
				return Convert.ToDouble(qa.AreaValue.Value);
			if (trovata is IIfcQuantityVolume qv)
				return Convert.ToDouble(qv.VolumeValue.Value);
			errore = "Computation method not implemented";
			return 0;
		}

		private string MatchName()
		{
			switch (LogicaDiCalcolo)
			{
				case TipoDiCosto.Area:
					return "NetSideArea";
				case TipoDiCosto.Volume:
					return "NetVolume";
			}
			return "";
		}

		internal static VoceDiListino Create(string logica, double cost)
		{
			VoceDiListino t = new VoceDiListino();
			if (logica == "Area")
				t.LogicaDiCalcolo = TipoDiCosto.Area;
			else if (logica == "Volume")
				t.LogicaDiCalcolo = TipoDiCosto.Volume;
			else
				t.LogicaDiCalcolo = TipoDiCosto.ACorpo;
			t.Costo = cost;
			return t;
		}
	}
}
