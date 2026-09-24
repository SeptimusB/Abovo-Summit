using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using AC = Aspose.Cells;
using Gear = SpreadsheetGear;

// Trial-only valid-input PMCost/RespCost ports. Not production replacements.
// Rejects malformed/unsorted lookup vectors rather than imitating VBA On Error Resume Next.
internal static class TrialFunctions
{
    internal static void VerifyKnownAnswers()
    {
        var years=Enumerable.Range(1,40).Select(i=>(double)i).ToArray();
        var rates=years.Select(y=>y*0.02).ToArray();
        double pm=Cost(true,new double[]{10,100,5,1,5,40},rates,years);
        double resp=Cost(false,new double[]{10,100,1,5,40},rates,years);
        if(Math.Abs(pm-80)>0.000001||Math.Abs(resp-16)>0.000001)
            throw new InvalidOperationException("Trial UDF known answer differs.");
    }
    internal static double Cost(bool pm,double[] scalar,double[] rates,double[] years)
    {
        int offset=pm?1:0,year=Convert.ToInt32(scalar[0]),units=Convert.ToInt32(scalar[1]),first=Convert.ToInt32(scalar[2+offset]),last=Convert.ToInt32(scalar[3+offset]);
        double unitCost=pm?scalar[2]:1,annual=(double)units/(last-first+1),total=0;
        if(last<first)return 0;
        if(years.Length==0)throw new ArgumentException("Lookup geometry.");
        double previous=Double.NegativeInfinity;foreach(double value in years){if(Double.IsNaN(value))continue;if(value<previous)throw new ArgumentException("Unsorted lookup.");previous=value;}
        for(int i=year,j=0;i>=first&&j<last-first+1;i--,j++)
        {
            int index=-1;double key=i-first+1;
            for(int k=0;k<years.Length;k++){if(Double.IsNaN(years[k]))continue;if(years[k]>key)break;index=k;}
            if(index<0||index>=rates.Length)throw new ArgumentException("Missing lookup.");
            total+=pm?unitCost*annual*rates[index]:annual*rates[index];
        }
        if(Double.IsNaN(total)||Double.IsInfinity(total))throw new ArithmeticException();
        return total;
    }
    static double[] Flatten(object value,bool blankYear=false)
    {
        var area=value as AC.ReferredArea;if(area!=null)value=area.GetValues();
        var list=new List<double>();FlattenInto(value,list,blankYear);return list.ToArray();
    }
    static void FlattenInto(object value,List<double> list,bool blankYear)
    {
        var array=value as Array;
        if(array!=null){foreach(object v in array)FlattenInto(v,list,blankYear);return;}
        if(value==null||Object.Equals(value,"")){list.Add(blankYear?Double.NaN:0);return;}
        if(value is string||value is bool)throw new ArgumentException("Non-numeric lookup.");
        list.Add(Convert.ToDouble(value));
    }
    internal sealed class AsposeFunctions:AC.AbstractCalculationEngine
    {
        public long Calls,InvalidCalls;
        public override void Calculate(AC.CalculationData data)
        {
            bool pm=String.Equals(data.FunctionName,"PMCost",StringComparison.OrdinalIgnoreCase);
            if(!pm&&!String.Equals(data.FunctionName,"RespCost",StringComparison.OrdinalIgnoreCase))return;
            Interlocked.Increment(ref Calls);
            try
            {
                int n=pm?6:5;if(data.ParamCount!=n+2)throw new ArgumentException();
                var scalars=new double[n];for(int i=0;i<n;i++){var v=Flatten(data.GetParamValue(i));if(v.Length!=1)throw new ArgumentException("Non-scalar input.");scalars[i]=v[0];}
                data.CalculatedValue=Cost(pm,scalars,Flatten(data.GetParamValue(n)),Flatten(data.GetParamValue(n+1),true));
            }
            catch(Exception e) {if(Interlocked.Increment(ref InvalidCalls)<=3)Console.WriteLine("Aspose UDF adapter: "+e.GetType().Name+": "+e.Message);data.CalculatedValue=AC.ErrorCellValueType.Value;}
        }
    }
    internal sealed class GearFunction:Gear.CustomFunctions.Function
    {
        readonly bool pm;public long Calls,InvalidCalls;
        internal GearFunction(bool planned):base(planned?"PMCost":"RespCost",Gear.CustomFunctions.Volatility.Invariant,Gear.CustomFunctions.ValueType.Number){pm=planned;}
        public override void Evaluate(Gear.CustomFunctions.IArguments args,Gear.CustomFunctions.IValue result)
        {
            Interlocked.Increment(ref Calls);
            try
            {
                int n=pm?6:5;if(args.Count!=n+2)throw new ArgumentException();
                var scalars=new double[n];for(int i=0;i<n;i++)scalars[i]=args.GetNumber(i);
                result.Number=Cost(pm,scalars,Vector(args,n,false),Vector(args,n+1,true));
            }
            catch(Exception e) {if(Interlocked.Increment(ref InvalidCalls)<=3)Console.WriteLine("SpreadsheetGear UDF adapter: "+e.GetType().Name+": "+e.Message);result.Error=Gear.ValueError.Value;}
        }
        static double[] Vector(Gear.CustomFunctions.IArguments args,int index,bool blankYear)
        {
            int rows,cols;args.GetArrayDimensions(index,out rows,out cols);var values=new double[rows*cols];
            for(int r=0;r<rows;r++)for(int c=0;c<cols;c++){var v=args.GetArrayValue(index,r,c);if(!v.IsNumber&&!v.IsEmpty)throw new ArgumentException("Non-numeric lookup argument "+index+" at "+r+","+c+" type="+v.Type);values[r*cols+c]=v.IsEmpty?(blankYear?Double.NaN:0):v.Number;}
            return values;
        }
    }
}
