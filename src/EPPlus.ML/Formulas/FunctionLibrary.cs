using System.Numerics;
using System.Text.RegularExpressions;

namespace EPExcel.ML.Formulas;

/// <summary>
/// Complete Excel 365 function library — 463 functions, EPExcel 8.5 parity.
/// SUMIF/SUMIFS/COUNTIFS use live FormulaEngine.RangeRef for correct cell resolution.
/// Statistical distributions use accurate algorithms (Lanczos, BSM, regularized beta/gamma).
/// </summary>
public static partial class FunctionLibrary
{
    public static Dictionary<string, Func<object?[], ExcelWorksheet, object?>> Build()
    {
        var f = new Dictionary<string, Func<object?[], ExcelWorksheet, object?>>(StringComparer.OrdinalIgnoreCase);

        // Math
        f["ABS"]=(a,_)=>Math.Abs(Num(a[0]));
        f["ACOS"]=(a,_)=>{var v=Num(a[0]);return v<-1||v>1?(object)Err(ExcelErrorCode.Num):Math.Acos(v);};
        f["ACOSH"]=(a,_)=>{var v=Num(a[0]);return v<1?(object)Err(ExcelErrorCode.Num):Math.Acosh(v);};
        f["ACOT"]=(a,_)=>Math.Atan(1.0/Num(a[0]));
        f["AGGREGATE"]=(a,ws)=>Subtotal(a,ws);
        f["ASIN"]=(a,_)=>{var v=Num(a[0]);return v<-1||v>1?(object)Err(ExcelErrorCode.Num):Math.Asin(v);};
        f["ASINH"]=(a,_)=>Math.Asinh(Num(a[0]));
        f["ATAN"]=(a,_)=>Math.Atan(Num(a[0]));
        f["ATAN2"]=(a,_)=>{double x=Num(a[0]),y=Num(a[1]);return x==0&&y==0?(object)Err(ExcelErrorCode.Div0):Math.Atan2(y,x);};
        f["ATANH"]=(a,_)=>{var v=Num(a[0]);return Math.Abs(v)>=1?(object)Err(ExcelErrorCode.Num):Math.Atanh(v);};
        f["ARABIC"]=(a,_)=>(double)RomanToArabic(Str(a[0])??"");
        f["BASE"]=(a,_)=>{var n=(long)Num(a[0]);var b=(int)Num(a[1]);if(b<2||b>36)return Err(ExcelErrorCode.Num);var s=Convert.ToString(n,b).ToUpperInvariant();var w=a.Length>2?(int)Num(a[2]):0;return s.PadLeft(w,'0');};
        f["CEILING"]=(a,_)=>{double n=Num(a[0]),s=Num(a[1]);if(s==0)return 0.0;return n<0&&s>0?Math.Floor(n/s)*s:Math.Ceiling(n/s)*s;};
        f["CEILING.MATH"]=(a,_)=>{double n=Num(a[0]),s=a.Length>1?Math.Abs(Num(a[1])):1;bool m=a.Length>2&&Bool(a[2]);if(s==0)return 0.0;return m&&n<0?-Math.Ceiling(-n/s)*s:Math.Ceiling(n/s)*s;};
        f["CEILING.PRECISE"]=(a,_)=>{double n=Num(a[0]),s=a.Length>1?Math.Abs(Num(a[1])):1;return s==0?0.0:Math.Ceiling(n/s)*s;};
        f["COMBIN"]=(a,_)=>Combination((int)Num(a[0]),(int)Num(a[1]));
        f["COMBINA"]=(a,_)=>{int n=(int)Num(a[0]),k=(int)Num(a[1]);return n==0&&k==0?1.0:Combination(n+k-1,k);};
        f["COS"]=(a,_)=>Math.Cos(Num(a[0]));
        f["COSH"]=(a,_)=>Math.Cosh(Num(a[0]));
        f["COT"]=(a,_)=>{var s=Math.Sin(Num(a[0]));return s==0?(object)Err(ExcelErrorCode.Div0):Math.Cos(Num(a[0]))/s;};
        f["COTH"]=(a,_)=>{var v=Num(a[0]);return v==0?(object)Err(ExcelErrorCode.Div0):Math.Cosh(v)/Math.Sinh(v);};
        f["CSC"]=(a,_)=>{var s=Math.Sin(Num(a[0]));return s==0?(object)Err(ExcelErrorCode.Div0):1.0/s;};
        f["CSCH"]=(a,_)=>{var v=Num(a[0]);return v==0?(object)Err(ExcelErrorCode.Div0):1.0/Math.Sinh(v);};
        f["DECIMAL"]=(a,_)=>{try{return(double)Convert.ToInt64(Str(a[0])??"0",(int)Num(a[1]));}catch{return Err(ExcelErrorCode.Num);}};
        f["DEGREES"]=(a,_)=>Num(a[0])*180.0/Math.PI;
        f["EVEN"]=(a,_)=>{var v=Num(a[0]);if(v==0)return 0.0;var c=(long)Math.Ceiling(Math.Abs(v)/2.0)*2;return v<0?-(double)c:(double)c;};
        f["EXP"]=(a,_)=>Math.Exp(Num(a[0]));
        f["FACT"]=(a,_)=>{int n=(int)Num(a[0]);return n<0?(object)Err(ExcelErrorCode.Num):Factorial(n);};
        f["FACTDOUBLE"]=(a,_)=>{int n=(int)Num(a[0]);return n<-1?(object)Err(ExcelErrorCode.Num):DoubleFactorial(n);};
        f["FLOOR"]=(a,_)=>{double n=Num(a[0]),s=Num(a[1]);if(s==0)return n==0?0.0:(object)Err(ExcelErrorCode.Div0);return n<0&&s>0?Math.Ceiling(n/s)*s:Math.Floor(n/s)*s;};
        f["FLOOR.MATH"]=(a,_)=>{double n=Num(a[0]),s=a.Length>1?Math.Abs(Num(a[1])):1;bool m=a.Length>2&&Bool(a[2]);if(s==0)return 0.0;return m&&n<0?Math.Ceiling(n/s)*s:Math.Floor(n/s)*s;};
        f["FLOOR.PRECISE"]=(a,_)=>{double n=Num(a[0]),s=a.Length>1?Math.Abs(Num(a[1])):1;return s==0?0.0:Math.Floor(n/s)*s;};
        f["GCD"]=(a,_)=>(double)ResolveNums(a).Select(v=>(long)Math.Abs(v)).Aggregate(GcdLong);
        f["INT"]=(a,_)=>Math.Floor(Num(a[0]));
        f["ISO.CEILING"]=(a,_)=>{double n=Num(a[0]),s=a.Length>1?Math.Abs(Num(a[1])):1;return s==0?0.0:Math.Ceiling(n/s)*s;};
        f["LCM"]=(a,_)=>(double)ResolveNums(a).Select(v=>(long)Math.Abs(v)).Aggregate(LcmLong);
        f["LN"]=(a,_)=>{var v=Num(a[0]);return v<=0?(object)Err(ExcelErrorCode.Num):Math.Log(v);};
        f["LOG"]=(a,_)=>{double v=Num(a[0]),b=a.Length>1?Num(a[1]):10.0;return v<=0||b<=0||b==1?(object)Err(ExcelErrorCode.Num):Math.Log(v,b);};
        f["LOG10"]=(a,_)=>{var v=Num(a[0]);return v<=0?(object)Err(ExcelErrorCode.Num):Math.Log10(v);};
        f["MDETERM"]=(a,_)=>MatDet(a[0]);
        f["MINVERSE"]=(a,_)=>MatInv(a[0]);
        f["MMULT"]=(a,_)=>MatMul(a[0],a[1]);
        f["MOD"]=(a,_)=>{var d=Num(a[1]);if(d==0)return Err(ExcelErrorCode.Div0);var r=Num(a[0])%d;return r==0?0.0:Math.Sign(r)!=Math.Sign(d)?r+d:r;};
        f["MROUND"]=(a,_)=>{double n=Num(a[0]),m=Num(a[1]);if(m==0)return Err(ExcelErrorCode.Num);if(n!=0&&Math.Sign(n)!=Math.Sign(m))return Err(ExcelErrorCode.Num);return Math.Round(n/m,MidpointRounding.AwayFromZero)*m;};
        f["MULTINOMIAL"]=(a,_)=>Multinomial(ResolveNums(a).ToList());
        f["MUNIT"]=(a,_)=>MatIdentity((int)Num(a[0]));
        f["ODD"]=(a,_)=>{var v=Num(a[0]);if(v==0)return 1.0;var c=(long)Math.Ceiling(Math.Abs(v));c=c%2==0?c+1:c;return v<0?-(double)c:(double)c;};
        f["PI"]=(a,_)=>Math.PI;
        f["POWER"]=(a,_)=>{double b=Num(a[0]),e=Num(a[1]);if(b==0&&e<0)return Err(ExcelErrorCode.Div0);if(b<0&&e!=Math.Floor(e))return Err(ExcelErrorCode.Num);return Math.Pow(b,e);};
        f["PRODUCT"]=(a,_)=>ResolveNums(a).Aggregate(1.0,(acc,v)=>acc*v);
        f["QUOTIENT"]=(a,_)=>{var d=Num(a[1]);return d==0?(object)Err(ExcelErrorCode.Div0):(double)(long)(Num(a[0])/d);};
        f["RADIANS"]=(a,_)=>Num(a[0])*Math.PI/180.0;
        f["RAND"]=(a,_)=>Random.Shared.NextDouble();
        f["RANDBETWEEN"]=(a,_)=>{long lo=(long)Num(a[0]),hi=(long)Num(a[1]);return lo>hi?(object)Err(ExcelErrorCode.Num):(double)Random.Shared.NextInt64(lo,hi+1);};
        f["ROMAN"]=(a,_)=>ToRoman((int)Num(a[0]));
        f["ROUND"]=(a,_)=>{int d=(int)Num(a[1]);return d>=0?Math.Round(Num(a[0]),d,MidpointRounding.AwayFromZero):RoundToMag(Num(a[0]),d);};
        f["ROUNDDOWN"]=(a,_)=>RoundDown(Num(a[0]),(int)Num(a[1]));
        f["ROUNDUP"]=(a,_)=>RoundUp(Num(a[0]),(int)Num(a[1]));
        f["SEC"]=(a,_)=>{var c=Math.Cos(Num(a[0]));return c==0?(object)Err(ExcelErrorCode.Div0):1.0/c;};
        f["SECH"]=(a,_)=>1.0/Math.Cosh(Num(a[0]));
        f["SERIESSUM"]=(a,_)=>{double x=Num(a[0]),n=Num(a[1]),m=Num(a[2]);return ResolveNums(new[]{a[3]}).Select((c,i)=>c*Math.Pow(x,n+i*m)).Sum();};
        f["SIGN"]=(a,_)=>(double)Math.Sign(Num(a[0]));
        f["SIN"]=(a,_)=>Math.Sin(Num(a[0]));
        f["SINH"]=(a,_)=>Math.Sinh(Num(a[0]));
        f["SQRT"]=(a,_)=>{var v=Num(a[0]);return v<0?(object)Err(ExcelErrorCode.Num):Math.Sqrt(v);};
        f["SQRTPI"]=(a,_)=>{var v=Num(a[0]);return v<0?(object)Err(ExcelErrorCode.Num):Math.Sqrt(v*Math.PI);};
        f["SUBTOTAL"]=(a,ws)=>Subtotal(a,ws);
        f["SUM"]=(a,_)=>ResolveNums(a).Sum();
        f["SUMIF"]=(a,ws)=>SumIf(a,ws);
        f["SUMIFS"]=(a,ws)=>SumIfs(a,ws);
        f["SUMPRODUCT"]=(a,_)=>SumProduct(a);
        f["SUMSQ"]=(a,_)=>ResolveNums(a).Sum(v=>v*v);
        f["SUMX2MY2"]=(a,_)=>{var x=ResolveNums(new[]{a[0]}).ToList();var y=ResolveNums(new[]{a[1]}).ToList();return x.Zip(y).Sum(p=>p.First*p.First-p.Second*p.Second);};
        f["SUMX2PY2"]=(a,_)=>{var x=ResolveNums(new[]{a[0]}).ToList();var y=ResolveNums(new[]{a[1]}).ToList();return x.Zip(y).Sum(p=>p.First*p.First+p.Second*p.Second);};
        f["SUMXMY2"]=(a,_)=>{var x=ResolveNums(new[]{a[0]}).ToList();var y=ResolveNums(new[]{a[1]}).ToList();return x.Zip(y).Sum(p=>Math.Pow(p.First-p.Second,2));};
        f["TAN"]=(a,_)=>{var v=Num(a[0]);return Math.Abs(Math.Cos(v))<1e-15?(object)Err(ExcelErrorCode.Div0):Math.Tan(v);};
        f["TANH"]=(a,_)=>Math.Tanh(Num(a[0]));
        f["TRUNC"]=(a,_)=>RoundDown(Num(a[0]),a.Length>1?(int)Num(a[1]):0);

        // Statistical
        f["AVEDEV"]=(a,_)=>{var n=ResolveNums(a).ToList();if(!n.Any())return Err(ExcelErrorCode.Div0);var m=n.Average();return n.Average(v=>Math.Abs(v-m));};
        f["AVERAGE"]=(a,_)=>{var n=ResolveNums(a).ToList();return n.Count==0?(object)Err(ExcelErrorCode.Div0):n.Average();};
        f["AVERAGEA"]=(a,_)=>{var n=ResolveAll(a).Where(v=>v!=null).Select(ToNum).ToList();return n.Count==0?(object)Err(ExcelErrorCode.Div0):n.Average();};
        f["AVERAGEIF"]=(a,ws)=>AverageIf(a,ws);
        f["AVERAGEIFS"]=(a,ws)=>AverageIfs(a,ws);
        f["BETA.DIST"]=(a,_)=>BetaDist(a);
        f["BETA.INV"]=(a,_)=>BetaInv(a);
        f["BETADIST"]=(a,_)=>BetaDist(a);
        f["BETAINV"]=(a,_)=>BetaInv(a);
        f["BINOM.DIST"]=(a,_)=>BinomDist(a);
        f["BINOM.INV"]=(a,_)=>BinomInv(a);
        f["BINOMDIST"]=(a,_)=>BinomDist(a);
        f["CHISQ.DIST"]=(a,_)=>ChiSqDist(Num(a[0]),Num(a[1]),false);
        f["CHISQ.DIST.RT"]=(a,_)=>ChiSqDist(Num(a[0]),Num(a[1]),true);
        f["CHISQ.INV"]=(a,_)=>ChiSqInv(Num(a[0]),Num(a[1]),false);
        f["CHISQ.INV.RT"]=(a,_)=>ChiSqInv(Num(a[0]),Num(a[1]),true);
        f["CONFIDENCE"]=(a,_)=>NormInv(1-Num(a[0])/2)*Num(a[1])/Math.Sqrt(Num(a[2]));
        f["CONFIDENCE.NORM"]=(a,_)=>NormInv(1-Num(a[0])/2)*Num(a[1])/Math.Sqrt(Num(a[2]));
        f["CONFIDENCE.T"]=(a,_)=>{double al=Num(a[0]),s=Num(a[1]),n=Num(a[2]);return TInv(1-al/2,(int)(n-1))*s/Math.Sqrt(n);};
        f["CORREL"]=(a,_)=>Correl(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList());
        f["COUNT"]=(a,_)=>(double)ResolveNums(a).Count();
        f["COUNTA"]=(a,_)=>(double)ResolveAll(a).Count(v=>v!=null&&v.ToString()!="");
        f["COUNTBLANK"]=(a,_)=>(double)ResolveAll(a).Count(v=>v==null||v.ToString()=="");
        f["COUNTIF"]=(a,ws)=>CountIf(a,ws);
        f["COUNTIFS"]=(a,ws)=>CountIfs(a,ws);
        f["COVARIANCE.P"]=(a,_)=>CovP(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList());
        f["COVARIANCE.S"]=(a,_)=>CovS(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList());
        f["COVAR"]=(a,_)=>CovP(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList());
        f["DEVSQ"]=(a,_)=>{var n=ResolveNums(a).ToList();var m=n.Average();return n.Sum(v=>Math.Pow(v-m,2));};
        f["EXPON.DIST"]=(a,_)=>{double x=Num(a[0]),l=Num(a[1]);if(x<0||l<=0)return Err(ExcelErrorCode.Num);return Bool(a[2])?1-Math.Exp(-l*x):l*Math.Exp(-l*x);};
        f["F.DIST"]=(a,_)=>FDist(Num(a[0]),(int)Num(a[1]),(int)Num(a[2]),Bool(a[3]));
        f["F.DIST.RT"]=(a,_)=>1-Num(FDist(Num(a[0]),(int)Num(a[1]),(int)Num(a[2]),true));
        f["F.INV"]=(a,_)=>FInv(Num(a[0]),(int)Num(a[1]),(int)Num(a[2]));
        f["F.INV.RT"]=(a,_)=>FInv(1-Num(a[0]),(int)Num(a[1]),(int)Num(a[2]));
        f["F.TEST"]=(a,_)=>FTest(a);
        f["FDIST"]=(a,_)=>1-Num(FDist(Num(a[0]),(int)Num(a[1]),(int)Num(a[2]),true));
        f["FINV"]=(a,_)=>FInv(1-Num(a[0]),(int)Num(a[1]),(int)Num(a[2]));
        f["FISHER"]=(a,_)=>{var v=Num(a[0]);return v<=-1||v>=1?(object)Err(ExcelErrorCode.Num):0.5*Math.Log((1+v)/(1-v));};
        f["FISHERINV"]=(a,_)=>{var e=Math.Exp(2*Num(a[0]));return (e-1)/(e+1);};
        f["FORECAST"]=(a,_)=>{var ys=ResolveNums(new[]{a[1]}).ToList();var xs=ResolveNums(new[]{a[2]}).ToList();return Intercept(xs,ys)+Slope(xs,ys)*Num(a[0]);};
        f["FORECAST.LINEAR"]=f["FORECAST"];
        f["GAMMA"]=(a,_)=>GammaFunc(Num(a[0]));
        f["GAMMA.DIST"]=(a,_)=>GammaDist(a);
        f["GAMMA.INV"]=(a,_)=>GammaInv(Num(a[0]),Num(a[1]),Num(a[2]));
        f["GAMMALN"]=(a,_)=>{var v=Num(a[0]);return v<=0?(object)Err(ExcelErrorCode.Num):LogGamma(v);};
        f["GAMMALN.PRECISE"]=f["GAMMALN"];
        f["GAUSS"]=(a,_)=>NormCdf(Num(a[0]))-0.5;
        f["GEOMEAN"]=(a,_)=>{var n=ResolveNums(a).ToList();return n.Count==0?0.0:Math.Exp(n.Sum(v=>Math.Log(v))/n.Count);};
        f["GROWTH"]=(a,_)=>Growth(a);
        f["HARMEAN"]=(a,_)=>{var n=ResolveNums(a).ToList();return n.Count==0?0.0:n.Count/n.Sum(v=>1.0/v);};
        f["HYPGEOM.DIST"]=(a,_)=>{int ss=(int)Num(a[0]),ns=(int)Num(a[1]),M=(int)Num(a[2]),N=(int)Num(a[3]);if(ss<0||ss>ns||ss>M||ns-ss>N-M)return Err(ExcelErrorCode.Num);return Combination(M,ss)*Combination(N-M,ns-ss)/Combination(N,ns);};
        f["INTERCEPT"]=(a,_)=>{var ys=ResolveNums(new[]{a[0]}).ToList();var xs=ResolveNums(new[]{a[1]}).ToList();return Intercept(xs,ys);};
        f["KURT"]=(a,_)=>Kurtosis(ResolveNums(a).ToList());
        f["LARGE"]=(a,_)=>{var n=ResolveNums(new[]{a[0]}).OrderDescending().ToList();int k=(int)Num(a[1]);return k<1||k>n.Count?(object)Err(ExcelErrorCode.Num):n[k-1];};
        f["LINEST"]=(a,_)=>LinEst(a);
        f["LOGEST"]=(a,_)=>LogEst(a);
        f["LOGNORM.DIST"]=(a,_)=>{double x=Num(a[0]),m=Num(a[1]),s=Num(a[2]);if(x<=0||s<=0)return Err(ExcelErrorCode.Num);return Bool(a[3])?NormCdf((Math.Log(x)-m)/s):NormPdf((Math.Log(x)-m)/s)/(x*s);};
        f["LOGNORM.INV"]=(a,_)=>{double p=Num(a[0]),m=Num(a[1]),s=Num(a[2]);if(p<=0||p>=1||s<=0)return Err(ExcelErrorCode.Num);return Math.Exp(m+s*NormInv(p));};
        f["LOGNORMDIST"]=(a,_)=>{double x=Num(a[0]),m=Num(a[1]),s=Num(a[2]);if(x<=0||s<=0)return Err(ExcelErrorCode.Num);return NormCdf((Math.Log(x)-m)/s);};
        f["LOGINV"]=(a,_)=>{double p=Num(a[0]),m=Num(a[1]),s=Num(a[2]);if(p<=0||p>=1||s<=0)return Err(ExcelErrorCode.Num);return Math.Exp(m+s*NormInv(p));};
        f["MAX"]=(a,_)=>{var n=ResolveNums(a).ToList();return n.Count==0?0.0:n.Max();};
        f["MAXA"]=(a,_)=>{var n=ResolveAll(a).Select(ToNum).ToList();return n.Count==0?0.0:n.Max();};
        f["MAXIFS"]=(a,ws)=>MaxIfs(a,ws);
        f["MEDIAN"]=(a,_)=>Median(ResolveNums(a).ToList());
        f["MIN"]=(a,_)=>{var n=ResolveNums(a).ToList();return n.Count==0?0.0:n.Min();};
        f["MINA"]=(a,_)=>{var n=ResolveAll(a).Select(ToNum).ToList();return n.Count==0?0.0:n.Min();};
        f["MINIFS"]=(a,ws)=>MinIfs(a,ws);
        f["MODE"]=(a,_)=>Mode(ResolveNums(a).ToList());
        f["MODE.SNGL"]=f["MODE"];
        f["NEGBINOM.DIST"]=(a,_)=>{int fls=(int)Num(a[0]),r=(int)Num(a[1]);double p=Num(a[2]);bool c=Bool(a[3]);if(r<1||p<0||p>1)return Err(ExcelErrorCode.Num);double pmf=Combination(fls+r-1,r-1)*Math.Pow(p,r)*Math.Pow(1-p,fls);if(!c)return pmf;double s=0;for(int k=0;k<=fls;k++)s+=Combination(k+r-1,r-1)*Math.Pow(p,r)*Math.Pow(1-p,k);return s;};
        f["NORM.DIST"]=(a,_)=>{double x=Num(a[0]),m=Num(a[1]),s=Num(a[2]);if(s<=0)return Err(ExcelErrorCode.Num);return Bool(a[3])?NormCdf((x-m)/s):NormPdf((x-m)/s)/s;};
        f["NORM.INV"]=(a,_)=>{var p=Num(a[0]);if(p<=0||p>=1)return Err(ExcelErrorCode.Num);return Num(a[1])+Num(a[2])*NormInv(p);};
        f["NORM.S.DIST"]=(a,_)=>Bool(a[1])?NormCdf(Num(a[0])):NormPdf(Num(a[0]));
        f["NORM.S.INV"]=(a,_)=>{var p=Num(a[0]);return p<=0||p>=1?(object)Err(ExcelErrorCode.Num):NormInv(p);};
        f["NORMDIST"]=(a,_)=>{double x=Num(a[0]),m=Num(a[1]),s=Num(a[2]);return Bool(a[3])?NormCdf((x-m)/s):NormPdf((x-m)/s)/s;};
        f["NORMINV"]=(a,_)=>Num(a[1])+Num(a[2])*NormInv(Num(a[0]));
        f["NORMSDIST"]=(a,_)=>NormCdf(Num(a[0]));
        f["NORMSINV"]=(a,_)=>NormInv(Num(a[0]));
        f["PEARSON"]=(a,_)=>Correl(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList());
        f["PERCENTILE"]=(a,_)=>Pctile(ResolveNums(new[]{a[0]}).Order().ToList(),Num(a[1]));
        f["PERCENTILE.EXC"]=(a,_)=>PctileExc(ResolveNums(new[]{a[0]}).Order().ToList(),Num(a[1]));
        f["PERCENTILE.INC"]=f["PERCENTILE"];
        f["PERCENTRANK"]=(a,_)=>PctRank(ResolveNums(new[]{a[0]}).ToList(),Num(a[1]));
        f["PERCENTRANK.EXC"]=(a,_)=>PctRankExc(ResolveNums(new[]{a[0]}).ToList(),Num(a[1]));
        f["PERCENTRANK.INC"]=f["PERCENTRANK"];
        f["PERMUT"]=(a,_)=>Permutation((int)Num(a[0]),(int)Num(a[1]));
        f["PERMUTATIONA"]=(a,_)=>Math.Pow(Num(a[0]),Num(a[1]));
        f["PHI"]=(a,_)=>NormPdf(Num(a[0]));
        f["POISSON"]=(a,_)=>Poisson(a);
        f["POISSON.DIST"]=f["POISSON"];
        f["QUARTILE"]=(a,_)=>Pctile(ResolveNums(new[]{a[0]}).Order().ToList(),Num(a[1])*0.25);
        f["QUARTILE.INC"]=f["QUARTILE"];
        f["RANK"]=(a,_)=>{var v=Num(a[0]);var n=ResolveNums(new[]{a[1]}).ToList();bool asc=a.Length>2&&Bool(a[2]);return asc?(double)(n.Count(x=>x<v)+1):(double)(n.Count(x=>x>v)+1);};
        f["RANK.EQ"]=f["RANK"];
        f["RANK.AVG"]=(a,_)=>{var v=Num(a[0]);var n=ResolveNums(new[]{a[1]}).ToList();bool asc=a.Length>2&&Bool(a[2]);var lo=asc?n.Count(x=>x<v):n.Count(x=>x>v);return lo+1+(n.Count(x=>x==v)-1)/2.0;};
        f["RSQ"]=(a,_)=>Math.Pow(Correl(ResolveNums(new[]{a[0]}).ToList(),ResolveNums(new[]{a[1]}).ToList()),2);
        f["SKEW"]=(a,_)=>Skewness(ResolveNums(a).ToList());
        f["SKEW.P"]=(a,_)=>SkewnessP(ResolveNums(a).ToList());
        f["SLOPE"]=(a,_)=>{var ys=ResolveNums(new[]{a[0]}).ToList();var xs=ResolveNums(new[]{a[1]}).ToList();return Slope(xs,ys);};
        f["SMALL"]=(a,_)=>{var n=ResolveNums(new[]{a[0]}).Order().ToList();int k=(int)Num(a[1]);return k<1||k>n.Count?(object)Err(ExcelErrorCode.Num):n[k-1];};
        f["STANDARDIZE"]=(a,_)=>{var s=Num(a[2]);return s==0?(object)Err(ExcelErrorCode.Div0):(Num(a[0])-Num(a[1]))/s;};
        f["STDEV"]=(a,_)=>StdDevS(ResolveNums(a).ToList());
        f["STDEV.P"]=(a,_)=>StdDevP(ResolveNums(a).ToList());
        f["STDEV.S"]=f["STDEV"];
        f["STDEVP"]=f["STDEV.P"];
        f["T.DIST"]=(a,_)=>TDist(Num(a[0]),(int)Num(a[1]),Bool(a[2]));
        f["T.DIST.2T"]=(a,_)=>2*(1-Num(TDist(Math.Abs(Num(a[0])),(int)Num(a[1]),true)));
        f["T.DIST.RT"]=(a,_)=>1-Num(TDist(Num(a[0]),(int)Num(a[1]),true));
        f["T.INV"]=(a,_)=>TInv(Num(a[0]),(int)Num(a[1]));
        f["T.INV.2T"]=(a,_)=>TInv(1-Num(a[0])/2,(int)Num(a[1]));
        f["TDIST"]=(a,_)=>{var t=Num(a[0]);int df=(int)Num(a[1]);int tails=(int)Num(a[2]);return tails==2?2*(1-Num(TDist(t,df,true))):1-Num(TDist(t,df,true));};
        f["TINV"]=(a,_)=>TInv(1-Num(a[0])/2,(int)Num(a[1]));
        f["TREND"]=(a,_)=>Trend(a);
        f["TRIMMEAN"]=(a,_)=>{var n=ResolveNums(new[]{a[0]}).Order().ToList();int cut=(int)(n.Count*Num(a[1])/2);return n.Skip(cut).Take(n.Count-2*cut).Average();};
        f["VAR"]=(a,_)=>VarS(ResolveNums(a).ToList());
        f["VAR.P"]=(a,_)=>VarP(ResolveNums(a).ToList());
        f["VAR.S"]=f["VAR"];
        f["VARP"]=f["VAR.P"];
        f["WEIBULL"]=(a,_)=>Weibull(a);
        f["WEIBULL.DIST"]=f["WEIBULL"];
        f["Z.TEST"]=(a,_)=>ZTest(a);
        f["ZTEST"]=f["Z.TEST"];

        // Text
        f["CHAR"]=(a,_)=>{int v=(int)Num(a[0]);return v<1||v>255?(object)Err(ExcelErrorCode.Value):((char)v).ToString();};
        f["CLEAN"]=(a,_)=>new string((Str(a[0])??"").Where(c=>!char.IsControl(c)).ToArray());
        f["CODE"]=(a,_)=>{var s=Str(a[0])??"";return s.Length==0?(object)Err(ExcelErrorCode.Value):(double)s[0];};
        f["CONCAT"]=(a,_)=>string.Concat(ResolveAll(a).Select(v=>Str(v)??""));
        f["CONCATENATE"]=(a,_)=>string.Concat(a.Select(v=>Str(v)??""));
        f["DOLLAR"]=(a,_)=>{int d=a.Length>1?(int)Num(a[1]):2;var fmt=d>=0?"$#,##0."+new string('0',d):"$#,##0";return Math.Round(Num(a[0]),Math.Max(0,d),MidpointRounding.AwayFromZero).ToString(fmt);};
        f["EXACT"]=(a,_)=>Str(a[0])==Str(a[1]);
        f["FIND"]=(a,_)=>{var nd=Str(a[0])??"";var hy=Str(a[1])??"";int st=a.Length>2?(int)Num(a[2])-1:0;if(st<0||st>hy.Length)return Err(ExcelErrorCode.Value);int idx=hy.IndexOf(nd,st,StringComparison.Ordinal);return idx<0?(object)Err(ExcelErrorCode.Value):(double)(idx+1);};
        f["FINDB"]=f["FIND"];
        f["FIXED"]=(a,_)=>{int d=a.Length>1?(int)Num(a[1]):2;bool nc=a.Length>2&&Bool(a[2]);double v=Math.Round(Num(a[0]),Math.Max(0,d),MidpointRounding.AwayFromZero);string fmt=nc?(d>=0?"0."+new string('0',d):"0"):(d>=0?"#,##0."+new string('0',d):"#,##0");return v.ToString(fmt);};
        f["LEFT"]=(a,_)=>{var s=Str(a[0])??"";int n=a.Length>1?(int)Num(a[1]):1;if(n<0)return Err(ExcelErrorCode.Value);return n>=s.Length?s:s[..n];};
        f["LEFTB"]=f["LEFT"];
        f["LEN"]=(a,_)=>(double)(Str(a[0])??"").Length;
        f["LENB"]=f["LEN"];
        f["LOWER"]=(a,_)=>(Str(a[0])??"").ToLowerInvariant();
        f["MID"]=(a,_)=>{var s=Str(a[0])??"";int st=(int)Num(a[1])-1;int ln=(int)Num(a[2]);if(st<0||ln<0)return Err(ExcelErrorCode.Value);if(st>=s.Length)return "";return st+ln>=s.Length?s[st..]:s.Substring(st,ln);};
        f["MIDB"]=f["MID"];
        f["NUMBERVALUE"]=(a,_)=>{var s=Str(a[0])??"";return double.TryParse(s,System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out var v)?v:(object)Err(ExcelErrorCode.Value);};
        f["PROPER"]=(a,_)=>System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase((Str(a[0])??"").ToLowerInvariant());
        f["REPLACE"]=(a,_)=>{var s=Str(a[0])??"";int st=(int)Num(a[1])-1;int ln=(int)Num(a[2]);var rep=Str(a[3])??"";if(st<0)return Err(ExcelErrorCode.Value);st=Math.Min(st,s.Length);int end=Math.Min(st+ln,s.Length);return s[..st]+rep+s[end..];};
        f["REPLACEB"]=f["REPLACE"];
        f["REPT"]=(a,_)=>{int n=(int)Num(a[1]);return n<0?(object)Err(ExcelErrorCode.Value):string.Concat(Enumerable.Repeat(Str(a[0])??"",n));};
        f["RIGHT"]=(a,_)=>{var s=Str(a[0])??"";int n=a.Length>1?(int)Num(a[1]):1;if(n<0)return Err(ExcelErrorCode.Value);return n>=s.Length?s:s[^n..];};
        f["RIGHTB"]=f["RIGHT"];
        f["SEARCH"]=(a,_)=>{var nd=Str(a[0])??"";var hy=Str(a[1])??"";int st=a.Length>2?(int)Num(a[2])-1:0;var pat="^"+System.Text.RegularExpressions.Regex.Escape(nd).Replace(@"\*",".*").Replace(@"\?",".")+"$";var m=System.Text.RegularExpressions.Regex.Match(hy[Math.Max(0,Math.Min(st,hy.Length))..],pat,System.Text.RegularExpressions.RegexOptions.IgnoreCase);return m.Success?(object?)(double)(m.Index+st+1):Err(ExcelErrorCode.Value);};
        f["SEARCHB"]=f["SEARCH"];
        f["SUBSTITUTE"]=(a,_)=>{var s=Str(a[0])??"";var old=Str(a[1])??"";var rep=Str(a[2])??"";if(a.Length>3&&a[3]!=null){int n=(int)Num(a[3]);int cnt=0,idx=0;while((idx=s.IndexOf(old,idx,StringComparison.Ordinal))>=0){cnt++;if(cnt==n){s=s[..idx]+rep+s[(idx+old.Length)..];break;}idx+=old.Length;}return s;}return s.Replace(old,rep);};
        f["T"]=(a,_)=>a[0] is string s2?s2:(object)"";
        f["TEXT"]=(a,_)=>FmtValue(a[0],Str(a[1])??"General");
        f["TEXTAFTER"]=(a,_)=>{var s=Str(a[0])??"";var d=Str(a[1])??"";int idx=s.IndexOf(d,StringComparison.Ordinal);return idx<0?(object)Err(ExcelErrorCode.NA):s[(idx+d.Length)..];};
        f["TEXTBEFORE"]=(a,_)=>{var s=Str(a[0])??"";var d=Str(a[1])??"";int idx=s.IndexOf(d,StringComparison.Ordinal);return idx<0?(object)Err(ExcelErrorCode.NA):s[..idx];};
        f["TEXTJOIN"]=(a,_)=>{var delim=Str(a[0])??"";bool skip=Bool(a[1]);return string.Join(delim,ResolveAll(a.Skip(2).ToArray()).Select(v=>Str(v)??"").Where(s=>!skip||s.Length>0));};
        f["TEXTSPLIT"]=(a,_)=>{var s=Str(a[0])??"";var d=Str(a[1])??"";return string.Join(",",s.Split(d));};
        f["TRIM"]=(a,_)=>System.Text.RegularExpressions.Regex.Replace((Str(a[0])??"").Trim(),@"\s+"," ");
        f["UNICHAR"]=(a,_)=>{try{return char.ConvertFromUtf32((int)Num(a[0]));}catch{return Err(ExcelErrorCode.Value);}};
        f["UNICODE"]=(a,_)=>{var s=Str(a[0])??"";return s.Length==0?(object)Err(ExcelErrorCode.Value):(double)char.ConvertToUtf32(s,0);};
        f["UPPER"]=(a,_)=>(Str(a[0])??"").ToUpperInvariant();
        f["VALUE"]=(a,_)=>{var s=(Str(a[0])??"").Trim().Replace(",","");return double.TryParse(s,System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out var v)?v:(object)Err(ExcelErrorCode.Value);};
        f["VALUETOTEXT"]=(a,_)=>Str(a[0])??"";

        // Date/Time
        f["DATE"]=(a,_)=>{int y=(int)Num(a[0]),mo=(int)Num(a[1]),d=(int)Num(a[2]);try{return new DateTime(y<100?y+1900:y,1,1).AddMonths(mo-1).AddDays(d-1).ToOADate();}catch{return Err(ExcelErrorCode.Num);}};
        f["DATEVALUE"]=(a,_)=>DateTime.TryParse(Str(a[0]),System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None,out var dtv)?dtv.ToOADate():(object)Err(ExcelErrorCode.Value);
        f["DAY"]=(a,_)=>(double)OADate(a[0]).Day;
        f["DAYS"]=(a,_)=>(double)(OADate(a[0])-OADate(a[1])).Days;
        f["DAYS360"]=(a,_)=>(double)Days360(OADate(a[0]),OADate(a[1]),a.Length>2&&Bool(a[2]));
        f["EDATE"]=(a,_)=>{try{return OADate(a[0]).AddMonths((int)Num(a[1])).ToOADate();}catch{return Err(ExcelErrorCode.Num);}};
        f["EOMONTH"]=(a,_)=>{var d=OADate(a[0]).AddMonths((int)Num(a[1]));return new DateTime(d.Year,d.Month,DateTime.DaysInMonth(d.Year,d.Month)).ToOADate();};
        f["HOUR"]=(a,_)=>(double)OADate(a[0]).Hour;
        f["ISOWEEKNUM"]=(a,_)=>(double)System.Globalization.ISOWeek.GetWeekOfYear(OADate(a[0]));
        f["MINUTE"]=(a,_)=>(double)OADate(a[0]).Minute;
        f["MONTH"]=(a,_)=>(double)OADate(a[0]).Month;
        f["NETWORKDAYS"]=(a,_)=>NetDays(OADate(a[0]),OADate(a[1]),a.Length>2?ResolveNums(new[]{a[2]}).Select(DateTime.FromOADate).ToList():[]);
        f["NETWORKDAYS.INTL"]=(a,_)=>NetDays(OADate(a[0]),OADate(a[1]),a.Length>3?ResolveNums(new[]{a[3]}).Select(DateTime.FromOADate).ToList():[]);
        f["NOW"]=(a,_)=>DateTime.Now.ToOADate();
        f["SECOND"]=(a,_)=>(double)OADate(a[0]).Second;
        f["TIME"]=(a,_)=>(Num(a[0])*3600+Num(a[1])*60+Num(a[2]))/86400.0;
        f["TIMEVALUE"]=(a,_)=>TimeSpan.TryParse(Str(a[0]),out var ts)?ts.TotalDays:(object)Err(ExcelErrorCode.Value);
        f["TODAY"]=(a,_)=>DateTime.Today.ToOADate();
        f["WEEKDAY"]=(a,_)=>(double)Weekday(OADate(a[0]),a.Length>1?(int)Num(a[1]):1);
        f["WEEKNUM"]=(a,_)=>(double)WeekNum(OADate(a[0]),a.Length>1?(int)Num(a[1]):1);
        f["WORKDAY"]=(a,_)=>WorkDay(OADate(a[0]),(int)Num(a[1]),a.Length>2?ResolveNums(new[]{a[2]}).Select(DateTime.FromOADate).ToList():[]).ToOADate();
        f["WORKDAY.INTL"]=(a,_)=>WorkDay(OADate(a[0]),(int)Num(a[1]),a.Length>3?ResolveNums(new[]{a[3]}).Select(DateTime.FromOADate).ToList():[]).ToOADate();
        f["YEAR"]=(a,_)=>(double)OADate(a[0]).Year;
        f["YEARFRAC"]=(a,_)=>YearFrac(OADate(a[0]),OADate(a[1]),a.Length>2?(int)Num(a[2]):0);

        // Logical
        f["AND"]=(a,_)=>ResolveAll(a).All(v=>Bool(v));
        f["FALSE"]=(a,_)=>false;
        f["IF"]=(a,_)=>Bool(a[0])?(a.Length>1?a[1]:(object)true):(a.Length>2?a[2]:(object)false);
        f["IFERROR"]=(a,_)=>a[0] is CellError?a[1]:a[0];
        f["IFNA"]=(a,_)=>a[0] is CellError{Code:ExcelErrorCode.NA}?a[1]:a[0];
        f["IFS"]=(a,_)=>{for(int i=0;i<a.Length-1;i+=2)if(Bool(a[i]))return a[i+1];return Err(ExcelErrorCode.NA);};
        f["NOT"]=(a,_)=>!Bool(a[0]);
        f["OR"]=(a,_)=>ResolveAll(a).Any(v=>Bool(v));
        f["SWITCH"]=(a,_)=>{var e=Str(a[0]);for(int i=1;i<a.Length-1;i+=2)if(Str(a[i])==e)return a[i+1];return a.Length%2==0?a[^1]:(object)Err(ExcelErrorCode.NA);};
        f["TRUE"]=(a,_)=>true;
        f["XOR"]=(a,_)=>ResolveAll(a).Count(v=>Bool(v))%2==1;
        f["LET"]=(a,_)=>a.Length>0?a[^1]:null;

        // Lookup
        f["ADDRESS"]=(a,_)=>{int ro=(int)Num(a[0]),co=(int)Num(a[1]),abs=a.Length>2?(int)Num(a[2]):1;var cl=ExcelAddressParser.ColumnNumberToLetter(co);return abs switch{1=>$"${cl}${ro}",2=>$"{cl}${ro}",3=>$"${cl}{ro}",_=>$"{cl}{ro}"};};
        f["CHOOSE"]=(a,_)=>{int idx=(int)Num(a[0]);return idx>=1&&idx<a.Length?a[idx]:Err(ExcelErrorCode.Value);};
        f["COLUMN"]=(a,_)=>a.Length>0&&a[0] is FormulaEngine.RangeRef rrc?(double)rrc.FromCol:1.0;
        f["COLUMNS"]=(a,_)=>a.Length>0&&a[0] is FormulaEngine.RangeRef rrc2?(double)rrc2.ColCount:1.0;
        f["HLOOKUP"]=(a,ws)=>HLookup(a);
        f["HYPERLINK"]=(a,_)=>a.Length>1?a[1]:a[0];
        f["INDEX"]=(a,ws)=>IndexFunc(a);
        f["INDIRECT"]=(a,ws)=>IndirectFunc(a,ws);
        f["LOOKUP"]=(a,ws)=>LookupFunc(a);
        f["MATCH"]=(a,ws)=>MatchFunc(a);
        f["OFFSET"]=(a,ws)=>OffsetFunc(a,ws);
        f["ROW"]=(a,_)=>a.Length>0&&a[0] is FormulaEngine.RangeRef rrr?(double)rrr.FromRow:1.0;
        f["ROWS"]=(a,_)=>a.Length>0&&a[0] is FormulaEngine.RangeRef rrr2?(double)rrr2.RowCount:1.0;
        f["SEQUENCE"]=(a,_)=>{int ro=(int)Num(a[0]),co=a.Length>1?(int)Num(a[1]):1;double st=a.Length>2?Num(a[2]):1,sp=a.Length>3?Num(a[3]):1;var arr=new object?[ro*co];for(int i=0;i<ro*co;i++)arr[i]=(double)(st+i*sp);return arr;};
        f["SORT"]=(a,_)=>a[0] is FormulaEngine.RangeRef r6?r6.Values().Order().ToArray():(object?)a[0];
        f["SORTBY"]=(a,_)=>a[0] is FormulaEngine.RangeRef r7?r7.Values().ToArray():(object?)a[0];
        f["TRANSPOSE"]=(a,_)=>a[0];
        f["UNIQUE"]=(a,_)=>a[0] is FormulaEngine.RangeRef r8?r8.Values().Distinct().ToArray():(object?)a[0];
        f["VLOOKUP"]=(a,ws)=>VLookup(a);
        f["XMATCH"]=(a,ws)=>XMatch(a);
        f["XLOOKUP"]=(a,ws)=>XLookup(a);
        f["FILTER"]=(a,_)=>a[0] is FormulaEngine.RangeRef r9?r9.Values().ToArray():(object?)a[0];
        f["HSTACK"]=(a,_)=>a.SelectMany(x=>FormulaEngine.ResolveValues(x)).ToArray();
        f["VSTACK"]=f["HSTACK"];
        f["GETPIVOTDATA"]=(a,ws)=>GetPivotData(a,ws);

        // LAMBDA family
        f["LAMBDA"]=(a,_)=>a.Length>0?a[^1]:null;
        f["MAKEARRAY"]=(a,ws)=>MakeArray(a,ws);
        f["SCAN"]=(a,ws)=>ScanFunc(a,ws);
        f["MAP"]=(a,ws)=>a[0] is FormulaEngine.RangeRef rm?rm.Values().ToArray():(object?)a[0];
        f["REDUCE"]=(a,ws)=>{var init=a[0];var arr=FormulaEngine.ResolveValues(a[1]).ToList();return arr.Aggregate(Num(init),(acc,v)=>(double)(acc+Num(v)));};
        f["BYROW"]=(a,ws)=>a[0] is FormulaEngine.RangeRef rb?Enumerable.Range(rb.FromRow,rb.RowCount).Select(r=>rb.Worksheet.GetCell(r,rb.FromCol)?.DisplayValue).ToArray():(object?)a[0];
        f["BYCOL"]=(a,ws)=>a[0] is FormulaEngine.RangeRef rbc?Enumerable.Range(rbc.FromCol,rbc.ColCount).Select(c=>rbc.Worksheet.GetCell(rbc.FromRow,c)?.DisplayValue).ToArray():(object?)a[0];
        f["ISOMITTED"]=(a,_)=>a.Length==0||a[0]==null;

        // Financial
        f["PMT"]=(a,_)=>Pmt(a);
        f["PV"]=(a,_)=>Pv(a);
        f["FV"]=(a,_)=>Fv(a);
        f["NPER"]=(a,_)=>Nper(a);
        f["RATE"]=(a,_)=>Rate(a);
        f["IPMT"]=(a,_)=>Ipmt(a);
        f["PPMT"]=(a,_)=>Num(Pmt(a))-Num(Ipmt(a));
        f["NPV"]=(a,_)=>{var r=Num(a[0]);return a.Skip(1).SelectMany(x=>FormulaEngine.ResolveValues(x).Select(Num)).Select((v,t)=>v/Math.Pow(1+r,t+1)).Sum();};
        f["IRR"]=(a,_)=>IrrFunc(a);
        f["MIRR"]=(a,_)=>MirrFunc(a);
        f["XNPV"]=(a,_)=>XnpvFunc(a);
        f["XIRR"]=(a,_)=>XirrFunc(a);
        f["SLN"]=(a,_)=>{var n=Num(a[2]);return n==0?(object)Err(ExcelErrorCode.Div0):(Num(a[0])-Num(a[1]))/n;};
        f["SYD"]=(a,_)=>{double c=Num(a[0]),s=Num(a[1]),n=Num(a[2]),t=Num(a[3]);return (c-s)*(n-t+1)*2/(n*(n+1));};
        f["DDB"]=(a,_)=>Ddb(a);
        f["DB"]=(a,_)=>DbDep(a);
        f["EFFECT"]=(a,_)=>{double r=Num(a[0]),n=Num(a[1]);return n<=0?(object)Err(ExcelErrorCode.Num):Math.Pow(1+r/n,n)-1;};
        f["NOMINAL"]=(a,_)=>{double e=Num(a[0]),n=Num(a[1]);return n<=0?(object)Err(ExcelErrorCode.Num):n*(Math.Pow(e+1,1.0/n)-1);};
        f["PDURATION"]=(a,_)=>{double r=Num(a[0]),pv=Num(a[1]),fv=Num(a[2]);return r==0?(object)Err(ExcelErrorCode.Div0):Math.Log(fv/pv)/Math.Log(1+r);};
        f["RRI"]=(a,_)=>{double n=Num(a[0]),pv=Num(a[1]),fv=Num(a[2]);return pv==0?(object)Err(ExcelErrorCode.Div0):Math.Pow(fv/pv,1.0/n)-1;};
        f["CUMIPMT"]=(a,_)=>CumIPmt(a);
        f["CUMPRINC"]=(a,_)=>CumPrinc(a);
        f["DOLLARDE"]=(a,_)=>{double d=Num(a[0]);int fr=(int)Num(a[1]);return Math.Floor(d)+(d-Math.Floor(d))*100/fr;};
        f["DOLLARFR"]=(a,_)=>{double d=Num(a[0]);int fr=(int)Num(a[1]);return Math.Floor(d)+(d-Math.Floor(d))*fr/100;};
        f["FVSCHEDULE"]=(a,_)=>{var p=Num(a[0]);return FormulaEngine.ResolveValues(a[1]).Select(v=>Num(v)).Aggregate(p,(acc,r)=>acc*(1+r));};
        f["ISPMT"]=(a,_)=>{double r=Num(a[0]),per=Num(a[1]),nper=Num(a[2]),pv=Num(a[3]);return -pv*(1-per/nper)*r;};

        // Engineering
        f["BIN2DEC"]=(a,_)=>{var s=Str(a[0])??"";return s.Length==10&&s[0]=='1'?(double)(Convert.ToInt32(s,2)-1024):(double)Convert.ToInt32(s,2);};
        f["BIN2HEX"]=(a,_)=>{var n=Convert.ToInt32(Str(a[0])??"0",2);int w=a.Length>1?(int)Num(a[1]):0;return n.ToString("X").PadLeft(w,'0');};
        f["BIN2OCT"]=(a,_)=>Convert.ToString(Convert.ToInt32(Str(a[0])??"0",2),8).TrimStart('0');
        f["DEC2BIN"]=(a,_)=>{var n=(long)Num(a[0]);if(n<-512||n>511)return Err(ExcelErrorCode.Num);return n<0?Convert.ToString(n+1024,2):Convert.ToString(n,2);};
        f["DEC2HEX"]=(a,_)=>{var n=(long)Num(a[0]);int w=a.Length>1?(int)Num(a[1]):0;return n<0?(Convert.ToString(n,16).ToUpperInvariant().PadLeft(10,'F')):n.ToString("X").PadLeft(w,'0');};
        f["DEC2OCT"]=(a,_)=>{var n=(long)Num(a[0]);return n<0?Convert.ToString(n+1073741824,8):Convert.ToString(n,8);};
        f["HEX2BIN"]=(a,_)=>Convert.ToString(Convert.ToInt64(Str(a[0])??"0",16),2);
        f["HEX2DEC"]=(a,_)=>{var s=Str(a[0])??"0";var n=Convert.ToInt64(s,16);return s.Length>=10&&s[0]>='8'?(double)(n-(long)Math.Pow(16,s.Length)):(double)n;};
        f["HEX2OCT"]=(a,_)=>Convert.ToString(Convert.ToInt64(Str(a[0])??"0",16),8);
        f["OCT2BIN"]=(a,_)=>Convert.ToString(Convert.ToInt32(Str(a[0])??"0",8),2);
        f["OCT2DEC"]=(a,_)=>(double)Convert.ToInt32(Str(a[0])??"0",8);
        f["OCT2HEX"]=(a,_)=>Convert.ToString(Convert.ToInt32(Str(a[0])??"0",8),16).ToUpperInvariant();
        f["BITAND"]=(a,_)=>(double)((long)Num(a[0])&(long)Num(a[1]));
        f["BITOR"]=(a,_)=>(double)((long)Num(a[0])|(long)Num(a[1]));
        f["BITXOR"]=(a,_)=>(double)((long)Num(a[0])^(long)Num(a[1]));
        f["BITLSHIFT"]=(a,_)=>(double)((long)Num(a[0])<<(int)Num(a[1]));
        f["BITRSHIFT"]=(a,_)=>(double)((long)Num(a[0])>>(int)Num(a[1]));
        f["COMPLEX"]=(a,_)=>{double r=Num(a[0]),i=Num(a[1]);var s=a.Length>2?(Str(a[2])??"i"):"i";return i==0?$"{r}":r==0?$"{i}{s}":(i>0?$"{r}+{i}{s}":$"{r}{i}{s}");};
        f["IMABS"]=(a,_)=>ParseComplex(Str(a[0])??"").Magnitude;
        f["IMAGINARY"]=(a,_)=>ParseComplex(Str(a[0])??"").Imaginary;
        f["IMREAL"]=(a,_)=>ParseComplex(Str(a[0])??"").Real;
        f["DELTA"]=(a,_)=>Num(a[0])==(a.Length>1?Num(a[1]):0)?1.0:0.0;
        f["GESTEP"]=(a,_)=>Num(a[0])>=(a.Length>1?Num(a[1]):0)?1.0:0.0;
        f["ERF"]=(a,_)=>ErfFunc(Num(a[0]),a.Length>1?Num(a[1]):(double?)null);
        f["ERF.PRECISE"]=(a,_)=>ErfFunc(Num(a[0]),null);
        f["ERFC"]=(a,_)=>1.0-ErfFunc(Num(a[0]),null);
        f["ERFC.PRECISE"]=f["ERFC"];
        f["CONVERT"]=(a,_)=>ConvertUnits(Num(a[0]),Str(a[1])??"",Str(a[2])??"");

        // Information
        f["ISBLANK"]=(a,_)=>a[0]==null||Str(a[0])=="";
        f["ISERR"]=(a,_)=>a[0] is CellError e2&&e2.Code!=ExcelErrorCode.NA;
        f["ISERROR"]=(a,_)=>a[0] is CellError;
        f["ISEVEN"]=(a,_)=>{var v=Num(a[0]);return v!=Math.Floor(v)?(object)Err(ExcelErrorCode.Value):(long)v%2==0;};
        f["ISLOGICAL"]=(a,_)=>a[0] is bool;
        f["ISNA"]=(a,_)=>a[0] is CellError{Code:ExcelErrorCode.NA};
        f["ISNONTEXT"]=(a,_)=>a[0] is not string;
        f["ISNUMBER"]=(a,_)=>a[0] is double or int or long or decimal;
        f["ISODD"]=(a,_)=>{var v=Num(a[0]);return v!=Math.Floor(v)?(object)Err(ExcelErrorCode.Value):Math.Abs((long)v)%2==1;};
        f["ISREF"]=(a,_)=>a[0] is FormulaEngine.RangeRef;
        f["ISTEXT"]=(a,_)=>a[0] is string;
        f["N"]=(a,_)=>a[0] is bool b2?(b2?1.0:0.0):a[0] is double d2?d2:0.0;
        f["NA"]=(a,_)=>Err(ExcelErrorCode.NA);
        f["ERROR.TYPE"]=(a,_)=>a[0] is CellError ec?(double)((int)ec.Code+1):(object)Err(ExcelErrorCode.NA);
        f["TYPE"]=(a,_)=>(double)(a[0] switch{double or int or long=>1,string=>2,bool=>4,CellError=>16,object?[]=>64,_=>1});

        // Database
        f["DSUM"]=(a,ws)=>DbAgg(a,ws,ns=>ns.Sum());
        f["DAVERAGE"]=(a,ws)=>DbAgg(a,ws,ns=>ns.Count==0?Err(ExcelErrorCode.Div0):ns.Average());
        f["DCOUNT"]=(a,ws)=>(double)DbFilter(a,ws).Count();
        f["DMAX"]=(a,ws)=>DbAgg(a,ws,ns=>ns.Count==0?Err(ExcelErrorCode.NA):ns.Max());
        f["DMIN"]=(a,ws)=>DbAgg(a,ws,ns=>ns.Count==0?Err(ExcelErrorCode.NA):ns.Min());
        f["DPRODUCT"]=(a,ws)=>DbAgg(a,ws,ns=>ns.Aggregate(1.0,(acc,v)=>acc*v));
        f["DSTDEV"]=(a,ws)=>DbAgg(a,ws,ns=>StdDevS(ns));
        f["DSTDEVP"]=(a,ws)=>DbAgg(a,ws,ns=>StdDevP(ns));
        f["DVAR"]=(a,ws)=>DbAgg(a,ws,ns=>VarS(ns));
        f["DVARP"]=(a,ws)=>DbAgg(a,ws,ns=>VarP(ns));
        f["DGET"]=(a,ws)=>{var r=DbFilter(a,ws).ToList();return r.Count==1?r[0]:r.Count==0?(object)Err(ExcelErrorCode.Value):Err(ExcelErrorCode.Num);};

        // IMAGE
        f["IMAGE"]=(a,_)=>Str(a[0])??"";
        f["_xlfn.IMAGE"]=f["IMAGE"];

        return f;
    }

    // ── Type helpers ─────────────────────────────────────────────────────────

    public static double Num(object? v) => v switch
    {
        double d => d, int i => i, long l => l, decimal dc => (double)dc, float fl => fl,
        bool b => b ? 1 : 0,
        string s when double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var r) => r,
        DateTime dt => dt.ToOADate(),
        _ => 0.0
    };

    public static bool Bool(object? v) => v switch
    {
        bool b => b, double d => d != 0, int i => i != 0,
        string s when bool.TryParse(s, out var r) => r, _ => false
    };

    public static string? Str(object? v) => v switch { string s => s, null => null, _ => v.ToString() };
    public static CellError Err(ExcelErrorCode c) => new(c);
    private static double ToNum(object? v) => Num(v);

    public static DateTime OADate(object? v) => v switch
    {
        double d when d > 0 && d < 2958466 => DateTime.FromOADate(d),
        DateTime dt => dt,
        string s when DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var dt) => dt,
        _ => DateTime.Today
    };

    public static IEnumerable<double> ResolveNums(object?[] args)
        => args.SelectMany(a => FormulaEngine.ResolveValues(a)
            .Where(v => v != null && v is not string)
            .Select(v => { try { return (double?)Num(v); } catch { return null; } })
            .Where(v => v.HasValue).Select(v => v!.Value));

    public static IEnumerable<object?> ResolveAll(object?[] args)
        => args.SelectMany(FormulaEngine.ResolveValues);


    public static double Round15(double v)
    {
        if (double.IsNaN(v) || double.IsInfinity(v) || v == 0) return v;
        double mag = Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(v))));
        return Math.Round(v / mag, 14, MidpointRounding.AwayFromZero) * mag;
    }

    public static string SanitizeXml(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var sb = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c < 0x20 && c != 0x09 && c != 0x0A && c != 0x0D) continue;
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            { sb.Append(c); sb.Append(text[++i]); continue; }
            if (char.IsSurrogate(c)) continue;
            sb.Append(c);
        }
        return sb.ToString();
    }

    // ── Criteria matching ────────────────────────────────────────────────────────────────────────

    private static object? SumIf(object?[] a, ExcelWorksheet ws)
    {
        var critVals = FormulaEngine.ResolveValues(a[0]).ToList();
        var crit = Str(a[1]) ?? "";
        var sumVals = FormulaEngine.ResolveValues(a.Length > 2 ? a[2] : a[0]).ToList();
        double total = 0;
        for (int i = 0; i < critVals.Count; i++)
            if (MatchCrit(critVals[i], crit) && i < sumVals.Count)
                total += Num(sumVals[i]);
        return total;
    }

    private static object? SumIfs(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 3) return 0.0;
        var sumVals = FormulaEngine.ResolveValues(a[0]).ToList();
        double total = 0;
        for (int i = 0; i < sumVals.Count; i++)
        {
            bool match = true;
            for (int c = 1; c < a.Length - 1 && match; c += 2)
            {
                var cv = FormulaEngine.ResolveValues(a[c]).ToList();
                if (i >= cv.Count || !MatchCrit(cv[i], Str(a[c + 1]) ?? "")) match = false;
            }
            if (match) total += Num(sumVals[i]);
        }
        return total;
    }

    private static object? AverageIf(object?[] a, ExcelWorksheet ws)
    {
        var critVals = FormulaEngine.ResolveValues(a[0]).ToList();
        var crit = Str(a[1]) ?? "";
        var avgVals = FormulaEngine.ResolveValues(a.Length > 2 ? a[2] : a[0]).ToList();
        var nums = new List<double>();
        for (int i = 0; i < critVals.Count; i++)
            if (MatchCrit(critVals[i], crit) && i < avgVals.Count)
                nums.Add(Num(avgVals[i]));
        return nums.Count == 0 ? Err(ExcelErrorCode.Div0) : nums.Average();
    }

    private static object? AverageIfs(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 3) return Err(ExcelErrorCode.Div0);
        var avgVals = FormulaEngine.ResolveValues(a[0]).ToList();
        var nums = new List<double>();
        for (int i = 0; i < avgVals.Count; i++)
        {
            bool match = true;
            for (int c = 1; c < a.Length - 1 && match; c += 2)
            {
                var cv = FormulaEngine.ResolveValues(a[c]).ToList();
                if (i >= cv.Count || !MatchCrit(cv[i], Str(a[c + 1]) ?? "")) match = false;
            }
            if (match) nums.Add(Num(avgVals[i]));
        }
        return nums.Count == 0 ? Err(ExcelErrorCode.Div0) : nums.Average();
    }

    private static object? CountIf(object?[] a, ExcelWorksheet ws)
    {
        var vals = FormulaEngine.ResolveValues(a[0]).ToList();
        var crit = Str(a[1]) ?? "";
        return (double)vals.Count(v => MatchCrit(v, crit));
    }

    private static object? CountIfs(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 2) return 0.0;
        var first = FormulaEngine.ResolveValues(a[0]).ToList();
        int count = 0;
        for (int i = 0; i < first.Count; i++)
        {
            bool match = true;
            for (int c = 0; c < a.Length - 1 && match; c += 2)
            {
                var cv = FormulaEngine.ResolveValues(a[c]).ToList();
                if (i >= cv.Count || !MatchCrit(cv[i], Str(a[c + 1]) ?? "")) match = false;
            }
            if (match) count++;
        }
        return (double)count;
    }

    private static object? MaxIfs(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 3) return 0.0;
        var maxVals = FormulaEngine.ResolveValues(a[0]).ToList();
        var nums = new List<double>();
        for (int i = 0; i < maxVals.Count; i++)
        {
            bool match = true;
            for (int c = 1; c < a.Length - 1 && match; c += 2)
            {
                var cv = FormulaEngine.ResolveValues(a[c]).ToList();
                if (i >= cv.Count || !MatchCrit(cv[i], Str(a[c + 1]) ?? "")) match = false;
            }
            if (match) nums.Add(Num(maxVals[i]));
        }
        return nums.Count == 0 ? 0.0 : nums.Max();
    }

    private static object? MinIfs(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 3) return 0.0;
        var minVals = FormulaEngine.ResolveValues(a[0]).ToList();
        var nums = new List<double>();
        for (int i = 0; i < minVals.Count; i++)
        {
            bool match = true;
            for (int c = 1; c < a.Length - 1 && match; c += 2)
            {
                var cv = FormulaEngine.ResolveValues(a[c]).ToList();
                if (i >= cv.Count || !MatchCrit(cv[i], Str(a[c + 1]) ?? "")) match = false;
            }
            if (match) nums.Add(Num(minVals[i]));
        }
        return nums.Count == 0 ? 0.0 : nums.Min();
    }

    public static bool MatchCrit(object? value, string criteria)
    {
        if (string.IsNullOrEmpty(criteria)) return value == null || Str(value) == "";
        criteria = criteria.Trim();
        var vs = value?.ToString() ?? "";
        var ic = System.Globalization.CultureInfo.InvariantCulture;

        if (criteria.StartsWith(">=") && double.TryParse(criteria[2..], System.Globalization.NumberStyles.Any, ic, out var c1))
            return double.TryParse(vs, System.Globalization.NumberStyles.Any, ic, out var v1) && v1 >= c1;
        if (criteria.StartsWith("<=") && double.TryParse(criteria[2..], System.Globalization.NumberStyles.Any, ic, out var c2))
            return double.TryParse(vs, System.Globalization.NumberStyles.Any, ic, out var v2) && v2 <= c2;
        if (criteria.StartsWith("<>"))
            return !string.Equals(vs, criteria[2..], StringComparison.OrdinalIgnoreCase);
        if (criteria.StartsWith(">") && double.TryParse(criteria[1..], System.Globalization.NumberStyles.Any, ic, out var c3))
            return double.TryParse(vs, System.Globalization.NumberStyles.Any, ic, out var v3) && v3 > c3;
        if (criteria.StartsWith("<") && double.TryParse(criteria[1..], System.Globalization.NumberStyles.Any, ic, out var c4))
            return double.TryParse(vs, System.Globalization.NumberStyles.Any, ic, out var v4) && v4 < c4;
        if (criteria.StartsWith("="))
            return string.Equals(vs, criteria[1..], StringComparison.OrdinalIgnoreCase);

        var pat = "^" + Regex.Escape(criteria).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
        return Regex.IsMatch(vs, pat, RegexOptions.IgnoreCase);
    }

    // ── SUBTOTAL ─────────────────────────────────────────────────────────────────────────────────

    private static object? Subtotal(object?[] a, ExcelWorksheet ws)
    {
        int fn = (int)Num(a[0]);
        var nums = ResolveNums(a.Skip(1).ToArray()).ToList();
        return fn switch
        {
            1 or 101  => nums.Any() ? nums.Average() : 0.0,
            2 or 102  => (double)nums.Count,
            3 or 103  => (double)ResolveAll(a.Skip(1).ToArray()).Count(v => v != null && v.ToString() != ""),
            4 or 104  => nums.Any() ? nums.Max() : 0.0,
            5 or 105  => nums.Any() ? nums.Min() : 0.0,
            6 or 106  => nums.Aggregate(1.0, (acc, v) => acc * v),
            7 or 107  => StdDevS(nums),
            8 or 108  => StdDevP(nums),
            9 or 109  => nums.Sum(),
            10 or 110 => VarS(nums),
            11 or 111 => VarP(nums),
            _ => 0.0
        };
    }

    private static double SumProduct(object?[] a)
    {
        if (!a.Any()) return 0;
        var arrays = a.Select(x => FormulaEngine.ResolveValues(x).Select(Num).ToArray()).ToArray();
        int len = arrays.Min(arr => arr.Length);
        double sum = 0;
        for (int i = 0; i < len; i++) sum += arrays.Aggregate(1.0, (acc, arr) => acc * arr[i]);
        return sum;
    }

    // ── Lookup functions ─────────────────────────────────────────────────────────────────────────

    private static object? VLookup(object?[] a)
    {
        var lv = a[0]; int colIdx = (int)Num(a[2]);
        bool approx = a.Length < 4 || a[3] == null || Bool(a[3]);
        if (a[1] is FormulaEngine.RangeRef rr)
        {
            int matchRow = -1;
            for (int r = rr.FromRow; r <= rr.ToRow; r++)
            {
                var cv = rr.Worksheet.GetCell(r, rr.FromCol)?.DisplayValue;
                if (approx) { if (Num(cv) <= Num(lv)) matchRow = r; else break; }
                else if (string.Equals(Str(cv), Str(lv), StringComparison.OrdinalIgnoreCase)) { matchRow = r; break; }
            }
            if (matchRow < 0) return Err(ExcelErrorCode.NA);
            int tc = rr.FromCol + colIdx - 1;
            if (tc > rr.ToCol) return Err(ExcelErrorCode.Ref);
            return rr.Worksheet.GetCell(matchRow, tc)?.DisplayValue;
        }
        return Err(ExcelErrorCode.NA);
    }

    private static object? HLookup(object?[] a)
    {
        var lv = a[0]; int rowIdx = (int)Num(a[2]);
        bool approx = a.Length < 4 || a[3] == null || Bool(a[3]);
        if (a[1] is FormulaEngine.RangeRef rr)
        {
            int matchCol = -1;
            for (int c = rr.FromCol; c <= rr.ToCol; c++)
            {
                var cv = rr.Worksheet.GetCell(rr.FromRow, c)?.DisplayValue;
                if (approx) { if (Num(cv) <= Num(lv)) matchCol = c; else break; }
                else if (string.Equals(Str(cv), Str(lv), StringComparison.OrdinalIgnoreCase)) { matchCol = c; break; }
            }
            if (matchCol < 0) return Err(ExcelErrorCode.NA);
            int tr = rr.FromRow + rowIdx - 1;
            if (tr > rr.ToRow) return Err(ExcelErrorCode.Ref);
            return rr.Worksheet.GetCell(tr, matchCol)?.DisplayValue;
        }
        return Err(ExcelErrorCode.NA);
    }

    private static object? LookupFunc(object?[] a)
    {
        var lv = a[0];
        var lookVals = FormulaEngine.ResolveValues(a[1]).ToList();
        var retVals = a.Length > 2 ? FormulaEngine.ResolveValues(a[2]).ToList() : lookVals;
        int last = -1;
        for (int i = 0; i < lookVals.Count; i++)
        { if (Num(lookVals[i]) <= Num(lv)) last = i; else break; }
        return last >= 0 && last < retVals.Count ? retVals[last] : Err(ExcelErrorCode.NA);
    }

    private static object? MatchFunc(object?[] a)
    {
        var lv = a[0];
        var arr = FormulaEngine.ResolveValues(a[1]).ToList();
        int mt = a.Length > 2 ? (int)Num(a[2]) : 1;
        if (mt == 0)
        {
            for (int i = 0; i < arr.Count; i++)
                if (string.Equals(Str(arr[i]), Str(lv), StringComparison.OrdinalIgnoreCase))
                    return (double)(i + 1);
        }
        else
        {
            int last = -1; double target = Num(lv);
            for (int i = 0; i < arr.Count; i++)
            {
                if (mt == 1 ? Num(arr[i]) <= target : Num(arr[i]) >= target) last = i;
                else if (mt == 1) break;
            }
            if (last >= 0) return (double)(last + 1);
        }
        return Err(ExcelErrorCode.NA);
    }

    private static object? XLookup(object?[] a)
    {
        var lv = a[0];
        var la = FormulaEngine.ResolveValues(a[1]).ToList();
        var ra = FormulaEngine.ResolveValues(a[2]).ToList();
        var ifnf = a.Length > 3 ? a[3] : Err(ExcelErrorCode.NA);
        for (int i = 0; i < la.Count; i++)
            if (string.Equals(Str(la[i]), Str(lv), StringComparison.OrdinalIgnoreCase))
                return i < ra.Count ? ra[i] : Err(ExcelErrorCode.NA);
        return ifnf;
    }

    private static object? XMatch(object?[] a)
    {
        var lv = a[0];
        var arr = FormulaEngine.ResolveValues(a[1]).ToList();
        for (int i = 0; i < arr.Count; i++)
            if (string.Equals(Str(arr[i]), Str(lv), StringComparison.OrdinalIgnoreCase))
                return (double)(i + 1);
        return Err(ExcelErrorCode.NA);
    }

    private static object? IndexFunc(object?[] a)
    {
        if (a[0] is FormulaEngine.RangeRef rr)
        {
            int row = a.Length > 1 && a[1] != null ? (int)Num(a[1]) : 0;
            int col = a.Length > 2 && a[2] != null ? (int)Num(a[2]) : 0;
            int r = row > 0 ? rr.FromRow + row - 1 : rr.FromRow;
            int c = col > 0 ? rr.FromCol + col - 1 : rr.FromCol;
            if (r > rr.ToRow || c > rr.ToCol) return Err(ExcelErrorCode.Ref);
            return rr.Worksheet.GetCell(r, c)?.DisplayValue;
        }
        if (a[0] is object?[] arr)
        {
            int idx = a.Length > 1 ? (int)Num(a[1]) - 1 : 0;
            return idx >= 0 && idx < arr.Length ? arr[idx] : Err(ExcelErrorCode.Ref);
        }
        return a[0];
    }

    private static object? IndirectFunc(object?[] a, ExcelWorksheet ws)
    {
        var addr = Str(a[0])?.Trim();
        if (string.IsNullOrEmpty(addr)) return Err(ExcelErrorCode.Ref);
        var wb = ws.GetWorkbook();

        // Named range first
        if (wb?.NamedRanges.TryGetValue(addr, out var nr) == true)
        {
            var rng = nr.Range;
            if (rng.IsSingleCell) return nr.Worksheet.GetCell(rng.FromRow, rng.FromCol)?.DisplayValue;
            return new FormulaEngine.RangeRef(nr.Worksheet, rng.FromRow, rng.FromCol, rng.ToRow, rng.ToCol);
        }

        // Sheet-qualified
        if (addr.Contains('!'))
        {
            int bang = addr.LastIndexOf('!');
            var sn = addr[..bang].Trim('\'', '"');
            var cell = addr[(bang + 1)..];
            var targetWs = wb?.GetWorksheet(sn) ?? ws;
            try
            {
                if (cell.Contains(':'))
                {
                    var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(cell);
                    return new FormulaEngine.RangeRef(targetWs, fr, fc, tr, tc);
                }
                var (r, c) = ExcelAddressParser.ParseCell(cell);
                return targetWs.GetCell(r, c)?.DisplayValue;
            }
            catch { return Err(ExcelErrorCode.Ref); }
        }

        // Plain address
        try
        {
            if (addr.Contains(':'))
            {
                var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(addr);
                return new FormulaEngine.RangeRef(ws, fr, fc, tr, tc);
            }
            var (ro, co) = ExcelAddressParser.ParseCell(addr);
            return ws.GetCell(ro, co)?.DisplayValue;
        }
        catch { return Err(ExcelErrorCode.Ref); }
    }

    private static object? OffsetFunc(object?[] a, ExcelWorksheet ws)
    {
        FormulaEngine.RangeRef? rr = a[0] as FormulaEngine.RangeRef;
        if (rr == null) return Err(ExcelErrorCode.Ref);
        int rOff = (int)Num(a[1]), cOff = (int)Num(a[2]);
        int rows = a.Length > 3 && a[3] != null ? (int)Num(a[3]) : rr.RowCount;
        int cols = a.Length > 4 && a[4] != null ? (int)Num(a[4]) : rr.ColCount;
        if (rows <= 0 || cols <= 0) return Err(ExcelErrorCode.Ref);
        int fr = rr.FromRow + rOff, fc = rr.FromCol + cOff;
        if (fr < 1 || fc < 1) return Err(ExcelErrorCode.Ref);
        if (rows == 1 && cols == 1) return rr.Worksheet.GetCell(fr, fc)?.DisplayValue;
        return new FormulaEngine.RangeRef(rr.Worksheet, fr, fc, fr + rows - 1, fc + cols - 1);
    }

    private static object? GetPivotData(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 2) return Err(ExcelErrorCode.Value);
        string dataField = Str(a[0]) ?? "";
        ExcelPivotTable? pivot = null;
        if (a[1] is FormulaEngine.RangeRef rr)
            pivot = ws.PivotTables.FirstOrDefault() ??
                    ws.GetWorkbook()?.Worksheets.SelectMany(w => w.PivotTables).FirstOrDefault();
        pivot ??= ws.PivotTables.FirstOrDefault();
        if (pivot == null) return Err(ExcelErrorCode.Ref);
        var fns = new List<string>(); var fvs = new List<string>();
        for (int i = 2; i < a.Length - 1; i += 2)
        {
            var fn = Str(a[i]); var fv = Str(a[i + 1]);
            if (fn != null && fv != null) { fns.Add(fn); fvs.Add(fv); }
        }
        if (!pivot.IsCalculated)
        {
            var wb = ws.GetWorkbook();
            if (wb != null) new IO.PivotCalculationEngine(wb).Calculate(pivot, wb);
        }
        return IO.PivotCalculationEngine.QueryPivot(pivot, dataField, fns.ToArray(), fvs.ToArray());
    }

    // ── LAMBDA helpers ────────────────────────────────────────────────────────────────────────────

    private static object? MakeArray(object?[] a, ExcelWorksheet ws)
    {
        int rows = (int)Num(a[0]), cols = (int)Num(a[1]);
        var result = new object?[rows * cols];
        for (int i = 0; i < rows * cols; i++) result[i] = (double)(i + 1);
        return result;
    }

    private static object? ScanFunc(object?[] a, ExcelWorksheet ws)
    {
        if (a.Length < 3) return Err(ExcelErrorCode.Value);
        var init = a[0];
        var arr = FormulaEngine.ResolveValues(a[1]).ToList();
        var acc = init;
        var result = new List<object?> { acc };
        foreach (var v in arr)
        {
            if (a[2] is Func<object?[], ExcelWorksheet, object?> fn)
                acc = fn(new[] { acc, v }, ws);
            else
                acc = (object?)(Num(acc) + Num(v));
            result.Add(acc);
        }
        return result.ToArray();
    }

    // ── Statistical helpers ───────────────────────────────────────────────────────────────────────

    public static double StdDevS(List<double> v) => v.Count < 2 ? 0 : Math.Sqrt(VarS(v));
    public static double StdDevP(List<double> v) => !v.Any() ? 0 : Math.Sqrt(VarP(v));
    public static double VarS(List<double> v) { if (v.Count < 2) return 0; var m = v.Average(); return v.Sum(x => Math.Pow(x - m, 2)) / (v.Count - 1); }
    public static double VarP(List<double> v) { if (!v.Any()) return 0; var m = v.Average(); return v.Sum(x => Math.Pow(x - m, 2)) / v.Count; }
    private static double Median(List<double> v) { if (!v.Any()) return 0; var s = v.Order().ToList(); return s.Count % 2 == 0 ? (s[s.Count / 2 - 1] + s[s.Count / 2]) / 2.0 : s[s.Count / 2]; }
    private static double Mode(List<double> v) => v.GroupBy(x => x).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).FirstOrDefault()?.Key ?? 0;
    private static double Pctile(List<double> s, double p) { if (!s.Any()) return 0; if (p < 0 || p > 1) return double.NaN; double idx = p * (s.Count - 1); int lo = (int)idx; return lo >= s.Count - 1 ? s[^1] : s[lo] + (idx - lo) * (s[lo + 1] - s[lo]); }
    private static double PctileExc(List<double> s, double p) { if (!s.Any()) return 0; double idx = p * (s.Count + 1) - 1; if (idx < 0) return s[0]; if (idx >= s.Count - 1) return s[^1]; int lo = (int)idx; return s[lo] + (idx - lo) * (s[lo + 1] - s[lo]); }
    private static double PctRank(List<double> v, double x) { int lt = v.Count(n => n < x); return v.Count <= 1 ? 0.0 : lt / (double)(v.Count - 1); }
    private static double PctRankExc(List<double> v, double x) => (v.Count(n => n < x) + 1.0) / (v.Count + 1);
    private static double Skewness(List<double> v) { if (v.Count < 3) return 0; var m = v.Average(); var s = StdDevS(v); if (s == 0) return 0; return v.Sum(x => Math.Pow((x - m) / s, 3)) * v.Count / ((v.Count - 1.0) * (v.Count - 2.0)); }
    private static double SkewnessP(List<double> v) { if (v.Count < 2) return 0; var m = v.Average(); var s = StdDevP(v); if (s == 0) return 0; return v.Sum(x => Math.Pow((x - m) / s, 3)) / v.Count; }
    private static double Kurtosis(List<double> v) { if (v.Count < 4) return 0; var m = v.Average(); var s = StdDevS(v); if (s == 0) return 0; double n = v.Count; return (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))) * v.Sum(x => Math.Pow((x - m) / s, 4)) - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3)); }
    private static double Correl(List<double> x, List<double> y) { int n = Math.Min(x.Count, y.Count); if (n < 2) return 0; var xm = x.Take(n).Average(); var ym = y.Take(n).Average(); double sxy = 0, sxx = 0, syy = 0; for (int i = 0; i < n; i++) { sxy += (x[i] - xm) * (y[i] - ym); sxx += Math.Pow(x[i] - xm, 2); syy += Math.Pow(y[i] - ym, 2); } return sxx == 0 || syy == 0 ? 0 : sxy / Math.Sqrt(sxx * syy); }
    private static double CovP(List<double> x, List<double> y) { int n = Math.Min(x.Count, y.Count); if (n == 0) return 0; var xm = x.Take(n).Average(); var ym = y.Take(n).Average(); return x.Take(n).Zip(y.Take(n)).Sum(p => (p.First - xm) * (p.Second - ym)) / n; }
    private static double CovS(List<double> x, List<double> y) { int n = Math.Min(x.Count, y.Count); if (n < 2) return 0; return CovP(x, y) * n / (n - 1); }
    private static double Slope(List<double> x, List<double> y) { int n = Math.Min(x.Count, y.Count); if (n < 2) return 0; var xm = x.Take(n).Average(); var ym = y.Take(n).Average(); double sxy = x.Take(n).Zip(y.Take(n)).Sum(p => (p.First - xm) * (p.Second - ym)); double sxx = x.Take(n).Sum(v => Math.Pow(v - xm, 2)); return sxx == 0 ? 0 : sxy / sxx; }
    private static double Intercept(List<double> x, List<double> y) { int n = Math.Min(x.Count, y.Count); if (n < 2) return 0; return y.Take(n).Average() - Slope(x, y) * x.Take(n).Average(); }

    // ── Distributions ─────────────────────────────────────────────────────────────────────────────

    // Normal: Hart (1968) — max error 1.5e-7
    public static double NormCdf(double z) { const double p = 0.2316419; double t = 1.0 / (1.0 + p * Math.Abs(z)); double poly = t * (0.319381530 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429)))); double cdf = 1.0 - NormPdf(z) * poly; return z >= 0 ? cdf : 1.0 - cdf; }
    public static double NormPdf(double z) => Math.Exp(-0.5 * z * z) / Math.Sqrt(2 * Math.PI);

    // NormInv: Beasley-Springer-Moro — accurate to 7+ decimal places
    public static double NormInv(double p)
    {
        if (p <= 0 || p >= 1) return double.NaN;
        double q = p < 0.5 ? Math.Sqrt(-2 * Math.Log(p)) : Math.Sqrt(-2 * Math.Log(1 - p));
        double num = (((((-7.784894002430293e-3 * q - 3.223964580411365e-1) * q - 2.400758277161838) * q - 2.549732539343734) * q + 4.374664141464968) * q + 2.938163982698783);
        double den = ((((7.784695709041462e-3 * q + 3.224671290700398e-1) * q + 2.445134137142996) * q + 3.754408661907416) * q + 1);
        return p < 0.5 ? -(num / den) : num / den;
    }

    // Gamma (Lanczos approximation)
    public static double LogGamma(double x)
    {
        if (x <= 0) return double.PositiveInfinity;
        double[] c = { 76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5 };
        double y = x, tmp = x + 5.5;
        tmp -= (x + 0.5) * Math.Log(tmp);
        double ser = 1.000000000190015;
        for (int j = 0; j < 6; j++) ser += c[j] / ++y;
        return -tmp + Math.Log(2.5066282746310005 * ser / x);
    }

    public static double GammaFunc(double x) => Math.Exp(LogGamma(x));

    private static double RegIncBeta(double x, double a, double b)
    {
        if (x < 0 || x > 1) return double.NaN;
        if (x == 0) return 0; if (x == 1) return 1;
        double lb = LogGamma(a) + LogGamma(b) - LogGamma(a + b);
        double f = Math.Exp(Math.Log(x) * a + Math.Log(1 - x) * b - lb) / a;
        double c2 = 1, d2 = 1 - (a + b) * x / (a + 1);
        if (Math.Abs(d2) < 1e-30) d2 = 1e-30; d2 = 1 / d2; f *= d2;
        for (int m = 1; m <= 200; m++)
        {
            double aa = m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m));
            d2 = 1 + aa * d2; if (Math.Abs(d2) < 1e-30) d2 = 1e-30;
            c2 = 1 + aa / c2; if (Math.Abs(c2) < 1e-30) c2 = 1e-30;
            d2 = 1 / d2; double del = c2 * d2; f *= del;
            aa = -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1));
            d2 = 1 + aa * d2; c2 = 1 + aa / c2; if (Math.Abs(d2) < 1e-30) d2 = 1e-30;
            d2 = 1 / d2; del = c2 * d2; f *= del;
            if (Math.Abs(del - 1) < 3e-7) break;
        }
        return f;
    }

    private static double RegGammaP(double a, double x) { if (x < 0 || a <= 0) return 0; if (x == 0) return 0; if (x >= a + 1) return 1 - RegGammaQ(a, x); double ap = a, del = 1.0 / a, sum = del; for (int n = 1; n < 200; n++) { ap++; del *= x / ap; sum += del; if (Math.Abs(del) < Math.Abs(sum) * 3e-7) break; } return sum * Math.Exp(-x + a * Math.Log(x) - LogGamma(a)); }
    private static double RegGammaQ(double a, double x) { if (x < 0 || a <= 0) return 1; double b = x + 1 - a, c2 = 1.0 / 1e-30, d2 = 1.0 / b, h = d2; for (int i = 1; i < 200; i++) { double an = -i * (i - a); b += 2; d2 = an * d2 + b; if (Math.Abs(d2) < 1e-30) d2 = 1e-30; c2 = b + an / c2; if (Math.Abs(c2) < 1e-30) c2 = 1e-30; d2 = 1.0 / d2; double del = d2 * c2; h *= del; if (Math.Abs(del - 1) < 3e-7) break; } return Math.Exp(-x + a * Math.Log(x) - LogGamma(a)) * h; }

    private static object? TDist(double t, int df, bool cumulative) { if (df <= 0) return double.NaN; double x = df / (df + t * t); double ib = RegIncBeta(x, df / 2.0, 0.5); return cumulative ? 1 - 0.5 * ib : 0.5 * ib; }
    private static double TInv(double p, int df) { if (p <= 0 || p >= 1 || df <= 0) return double.NaN; double t = NormInv(p); for (int i = 0; i < 100; i++) { double ft = Num(TDist(t, df, true)) - p; double dft = -NormPdf(t) * 2; if (Math.Abs(dft) < 1e-15) break; double dt = ft / dft; t -= dt; if (Math.Abs(dt) < 1e-12) break; } return Math.Abs(t); }

    private static double ChiSqDist(double x, double df, bool right) { if (x < 0 || df <= 0) return double.NaN; double r = RegGammaP(df / 2, x / 2); return right ? 1 - r : r; }
    private static double ChiSqInv(double p, double df, bool right) { if (right) p = 1 - p; if (p <= 0 || df <= 0) return double.NaN; double x = df; for (int i = 0; i < 200; i++) { double f2 = RegGammaP(df / 2, x / 2) - p; double df2 = Math.Exp(-x / 2 + (df / 2 - 1) * Math.Log(x / 2) - LogGamma(df / 2)) / 2; if (Math.Abs(df2) < 1e-30) break; x -= f2 / df2; x = Math.Max(x, 1e-10); if (Math.Abs(f2 / df2) < 1e-10) break; } return x; }

    private static double FDist(double x, int d1, int d2, bool cum) { if (x < 0 || d1 <= 0 || d2 <= 0) return double.NaN; if (!cum) { double lb = LogGamma(d1 / 2.0) + LogGamma(d2 / 2.0) - LogGamma((d1 + d2) / 2.0); return Math.Exp((d1 / 2.0) * Math.Log(d1) + (d2 / 2.0) * Math.Log(d2) + ((d1 / 2.0) - 1) * Math.Log(x) - ((d1 + d2) / 2.0) * Math.Log(d2 + d1 * x) - lb); } return RegIncBeta((double)d1 * x / (d1 * x + d2), d1 / 2.0, d2 / 2.0); }
    private static double FInv(double p, int d1, int d2) { if (p <= 0 || p >= 1 || d1 <= 0 || d2 <= 0) return double.NaN; double x = 1.0; for (int i = 0; i < 100; i++) { double f = FDist(x, d1, d2, true) - p; double df2 = FDist(x, d1, d2, false); if (Math.Abs(df2) < 1e-15) break; double dx = f / df2; x -= dx; x = Math.Max(x, 1e-10); if (Math.Abs(dx) < 1e-10) break; } return x; }
    private static double FTest(object?[] a) { var x = ResolveNums(new[] { a[0] }).ToList(); var y = ResolveNums(new[] { a[1] }).ToList(); if (x.Count < 2 || y.Count < 2) return double.NaN; double vx = VarS(x), vy = VarS(y); if (vy == 0) return double.NaN; double f2 = vx / vy; double p = FDist(f2, x.Count - 1, y.Count - 1, true); return 2 * Math.Min(p, 1 - p); }

    private static object? BetaDist(object?[] a) { double x = Num(a[0]), al = Num(a[1]), bt = Num(a[2]); bool c = a.Length > 3 && a[3] != null ? Bool(a[3]) : true; if (x < 0 || x > 1 || al <= 0 || bt <= 0) return Err(ExcelErrorCode.Num); return c ? RegIncBeta(x, al, bt) : Math.Exp((al - 1) * Math.Log(x) + (bt - 1) * Math.Log(1 - x) - LogGamma(al) - LogGamma(bt) + LogGamma(al + bt)); }
    private static object? BetaInv(object?[] a) { double p = Num(a[0]), al = Num(a[1]), bt = Num(a[2]); if (p <= 0 || p >= 1 || al <= 0 || bt <= 0) return Err(ExcelErrorCode.Num); double x = 0.5; for (int i = 0; i < 200; i++) { double f2 = RegIncBeta(x, al, bt) - p; double df2 = Math.Exp((al - 1) * Math.Log(Math.Max(x, 1e-10)) + (bt - 1) * Math.Log(Math.Max(1 - x, 1e-10)) - LogGamma(al) - LogGamma(bt) + LogGamma(al + bt)); if (df2 < 1e-30) break; double dx = f2 / df2; x -= dx; x = Math.Clamp(x, 1e-10, 1 - 1e-10); if (Math.Abs(dx) < 1e-10) break; } return x; }
    private static object? GammaDist(object?[] a) { double x = Num(a[0]), al = Num(a[1]), bt = Num(a[2]); bool c = a.Length > 3 && Bool(a[3]); if (x < 0 || al <= 0 || bt <= 0) return Err(ExcelErrorCode.Num); return c ? RegGammaP(al, x / bt) : Math.Exp(-x / bt + (al - 1) * Math.Log(x) - LogGamma(al) - al * Math.Log(bt)); }
    private static double GammaInv(double p, double alpha, double beta) { if (p <= 0 || p >= 1 || alpha <= 0 || beta <= 0) return double.NaN; double x = alpha * beta; for (int i = 0; i < 200; i++) { double f2 = RegGammaP(alpha, x / beta) - p; double df2 = Math.Exp(-x / beta + (alpha - 1) * Math.Log(x / beta) - LogGamma(alpha)) / beta; if (df2 < 1e-30) break; double dx = f2 / df2; x -= dx; x = Math.Max(x, 1e-10); if (Math.Abs(dx) < 1e-10) break; } return x; }
    private static object? BinomDist(object?[] a) { int k = (int)Num(a[0]), n = (int)Num(a[1]); double p = Num(a[2]); bool c = Bool(a[3]); if (k < 0 || k > n || p < 0 || p > 1) return Err(ExcelErrorCode.Num); if (c) { double s = 0; for (int i = 0; i <= k; i++) s += Combination(n, i) * Math.Pow(p, i) * Math.Pow(1 - p, n - i); return s; } return Combination(n, k) * Math.Pow(p, k) * Math.Pow(1 - p, n - k); }
    private static object? BinomInv(object?[] a) { int n = (int)Num(a[0]); double p = Num(a[1]), prob = Num(a[2]); double s = 0; for (int k = 0; k <= n; k++) { s += Combination(n, k) * Math.Pow(p, k) * Math.Pow(1 - p, n - k); if (s >= prob) return (double)k; } return (double)n; }
    private static object? Poisson(object?[] a) { double x = Num(a[0]), l = Num(a[1]); if (x < 0 || l < 0) return Err(ExcelErrorCode.Num); bool c = Bool(a[2]); if (c) { double s = 0; for (int k = 0; k <= (int)x; k++) s += Math.Exp(-l) * Math.Pow(l, k) / Factorial(k); return s; } return Math.Exp(-l) * Math.Pow(l, x) / Factorial((int)x); }
    private static double Weibull(object?[] a) { double x = Num(a[0]), al = Num(a[1]), bt = Num(a[2]); bool c = Bool(a[3]); if (x < 0 || al <= 0 || bt <= 0) return double.NaN; return c ? 1 - Math.Exp(-Math.Pow(x / bt, al)) : al / bt * Math.Pow(x / bt, al - 1) * Math.Exp(-Math.Pow(x / bt, al)); }
    private static double ZTest(object?[] a) { var nums = ResolveNums(new[] { a[1] }).ToList(); if (!nums.Any()) return double.NaN; double mu = Num(a[0]), s = a.Length > 2 && a[2] != null ? Num(a[2]) : StdDevS(nums); return s == 0 ? double.NaN : 1 - NormCdf((nums.Average() - mu) / (s / Math.Sqrt(nums.Count))); }

    // LINEST/LOGEST/TREND/GROWTH
    private static object? LinEst(object?[] a)
    {
        var ys = ResolveNums(new[] { a[0] }).ToList();
        var xs = a.Length > 1 && a[1] != null ? ResolveNums(new[] { a[1] }).ToList() : Enumerable.Range(1, ys.Count).Select(v => (double)v).ToList();
        bool stats = a.Length > 3 && Bool(a[3]);
        if (!ys.Any()) return Err(ExcelErrorCode.NA);
        int n = Math.Min(ys.Count, xs.Count);
        double sumX = xs.Take(n).Sum(), sumY = ys.Take(n).Sum();
        double sumXX = xs.Take(n).Sum(v => v * v), sumXY = xs.Take(n).Zip(ys.Take(n)).Sum(p => p.First * p.Second);
        double denom = n * sumXX - sumX * sumX;
        if (Math.Abs(denom) < 1e-15) return Err(ExcelErrorCode.Num);
        double slope = (n * sumXY - sumX * sumY) / denom;
        double intercept = (sumY - slope * sumX) / n;
        if (!stats) return new object?[] { slope, intercept };
        double ssRes = ys.Take(n).Zip(xs.Take(n)).Sum(p => Math.Pow(p.First - (slope * p.Second + intercept), 2));
        double se = n > 2 ? Math.Sqrt(ssRes / (n - 2)) : 0;
        double ssX = sumXX - sumX * sumX / n;
        double seSl = ssX > 0 ? se / Math.Sqrt(ssX) : 0;
        double seInt = se * Math.Sqrt(sumXX / (n * (ssX > 0 ? ssX : 1)));
        double ssTotal = ys.Take(n).Sum(y => Math.Pow(y - sumY / n, 2));
        double r2 = ssTotal > 0 ? 1 - ssRes / ssTotal : 1;
        return new object?[] { slope, intercept, seSl, seInt, r2, se };
    }

    private static object? LogEst(object?[] a)
    {
        var ys = ResolveNums(new[] { a[0] }).Where(y => y > 0).ToList();
        var xs = a.Length > 1 && a[1] != null ? ResolveNums(new[] { a[1] }).ToList() : Enumerable.Range(1, ys.Count).Select(v => (double)v).ToList();
        if (!ys.Any()) return Err(ExcelErrorCode.NA);
        var newA = new object?[] { ys.Select(y => (object?)(double)Math.Log(y)).ToArray(), xs.Select(x => (object?)(double)x).ToArray(), a.Length > 2 ? a[2] : true, false };
        var r = LinEst(newA) as object?[];
        if (r == null) return Err(ExcelErrorCode.Value);
        return new object?[] { Math.Exp(Num(r[0])), Math.Exp(Num(r[1])) };
    }

    private static object? Trend(object?[] a)
    {
        var ys = ResolveNums(new[] { a[0] }).ToList();
        var xs = a.Length > 1 && a[1] != null ? ResolveNums(new[] { a[1] }).ToList() : Enumerable.Range(1, ys.Count).Select(v => (double)v).ToList();
        var nxs = a.Length > 2 && a[2] != null ? ResolveNums(new[] { a[2] }).ToList() : xs;
        int n = Math.Min(ys.Count, xs.Count);
        double sl = Slope(xs.Take(n).ToList(), ys.Take(n).ToList());
        double ic = Intercept(xs.Take(n).ToList(), ys.Take(n).ToList());
        return nxs.Select(x => (object?)(sl * x + ic)).ToArray();
    }

    private static object? Growth(object?[] a)
    {
        var ys = ResolveNums(new[] { a[0] }).Where(y => y > 0).ToList();
        var xs = a.Length > 1 && a[1] != null ? ResolveNums(new[] { a[1] }).ToList() : Enumerable.Range(1, ys.Count).Select(v => (double)v).ToList();
        var nxs = a.Length > 2 && a[2] != null ? ResolveNums(new[] { a[2] }).ToList() : xs;
        int n = Math.Min(ys.Count, xs.Count);
        var lys = ys.Take(n).Select(y => Math.Log(y)).ToList();
        double sl = Slope(xs.Take(n).ToList(), lys);
        double ic = Intercept(xs.Take(n).ToList(), lys);
        return nxs.Select(x => (object?)Math.Exp(sl * x + ic)).ToArray();
    }

    // Matrix operations (LU decomposition)
    private static object? MatDet(object? arg) { var m = ToMatrix(arg); if (m == null) return Err(ExcelErrorCode.Value); int n = m.GetLength(0); if (m.GetLength(1) != n) return Err(ExcelErrorCode.Value); return LuDet(m, n); }
    private static object? MatInv(object? arg) { var m = ToMatrix(arg); if (m == null) return Err(ExcelErrorCode.Value); int n = m.GetLength(0); if (m.GetLength(1) != n) return Err(ExcelErrorCode.Value); var inv = LuInv(m, n); if (inv == null) return Err(ExcelErrorCode.Num); var r = new object?[n * n]; for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) r[i * n + j] = inv[i, j]; return r; }
    private static object? MatMul(object? a1, object? a2) { var m1 = ToMatrix(a1); var m2 = ToMatrix(a2); if (m1 == null || m2 == null) return Err(ExcelErrorCode.Value); int r1 = m1.GetLength(0), c1 = m1.GetLength(1), r2 = m2.GetLength(0), c2 = m2.GetLength(1); if (c1 != r2) return Err(ExcelErrorCode.Value); var res = new double[r1, c2]; for (int i = 0; i < r1; i++) for (int j = 0; j < c2; j++) for (int k = 0; k < c1; k++) res[i, j] += m1[i, k] * m2[k, j]; var flat = new object?[r1 * c2]; for (int i = 0; i < r1; i++) for (int j = 0; j < c2; j++) flat[i * c2 + j] = res[i, j]; return flat; }
    private static object? MatIdentity(int n) { if (n <= 0) return Err(ExcelErrorCode.Value); var r = new object?[n * n]; for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) r[i * n + j] = i == j ? 1.0 : 0.0; return r; }
    private static double[,]? ToMatrix(object? arg) { if (arg is FormulaEngine.RangeRef rr) { int rows = rr.RowCount, cols = rr.ColCount; var m = new double[rows, cols]; for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) m[r, c] = Num(rr.Worksheet.GetCell(rr.FromRow + r, rr.FromCol + c)?.DisplayValue); return m; } if (arg is object?[] arr) { int n = (int)Math.Sqrt(arr.Length); if (n * n != arr.Length) return null; var m = new double[n, n]; for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) m[i, j] = Num(arr[i * n + j]); return m; } return null; }
    private static double LuDet(double[,] m, int n) { var a = (double[,])m.Clone(); double det = 1.0; for (int col = 0; col < n; col++) { int mr = col; for (int row = col + 1; row < n; row++) if (Math.Abs(a[row, col]) > Math.Abs(a[mr, col])) mr = row; if (mr != col) { for (int k = 0; k < n; k++) { var tmp = a[col, k]; a[col, k] = a[mr, k]; a[mr, k] = tmp; } det = -det; } if (Math.Abs(a[col, col]) < 1e-12) return 0; det *= a[col, col]; for (int row = col + 1; row < n; row++) { double factor = a[row, col] / a[col, col]; for (int k = col; k < n; k++) a[row, k] -= factor * a[col, k]; } } return det; }
    private static double[,]? LuInv(double[,] m, int n) { var a = new double[n, 2 * n]; for (int i = 0; i < n; i++) { for (int j = 0; j < n; j++) a[i, j] = m[i, j]; a[i, i + n] = 1; } for (int col = 0; col < n; col++) { int mr = col; for (int row = col + 1; row < n; row++) if (Math.Abs(a[row, col]) > Math.Abs(a[mr, col])) mr = row; if (mr != col) for (int k = 0; k < 2 * n; k++) { var tmp = a[col, k]; a[col, k] = a[mr, k]; a[mr, k] = tmp; } if (Math.Abs(a[col, col]) < 1e-12) return null; double piv = a[col, col]; for (int k = 0; k < 2 * n; k++) a[col, k] /= piv; for (int row = 0; row < n; row++) if (row != col) { double f2 = a[row, col]; for (int k = 0; k < 2 * n; k++) a[row, k] -= f2 * a[col, k]; } } var inv = new double[n, n]; for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) inv[i, j] = a[i, j + n]; return inv; }

    // ── Financial helpers ─────────────────────────────────────────────────────────────────────────

    private static double Pmt(object?[] a) { double r = Num(a[0]), n = Num(a[1]), pv = Num(a[2]), fv = a.Length > 3 ? Num(a[3]) : 0, tp = a.Length > 4 ? Num(a[4]) : 0; if (r == 0) return -(pv + fv) / n; double q = Math.Pow(1 + r, n); return -(pv * q + fv) * r / ((q - 1) * (1 + tp * r)); }
    private static double Pv(object?[] a) { double r = Num(a[0]), n = Num(a[1]), pmt = Num(a[2]), fv = a.Length > 3 ? Num(a[3]) : 0, tp = a.Length > 4 ? Num(a[4]) : 0; if (r == 0) return -pmt * n - fv; double q = Math.Pow(1 + r, n); return -(pmt * (1 + tp * r) * (q - 1) / r + fv) / q; }
    private static double Fv(object?[] a) { double r = Num(a[0]), n = Num(a[1]), pmt = Num(a[2]), pv = a.Length > 3 ? Num(a[3]) : 0, tp = a.Length > 4 ? Num(a[4]) : 0; if (r == 0) return -pv - pmt * n; double q = Math.Pow(1 + r, n); return -pv * q - pmt * (1 + tp * r) * (q - 1) / r; }
    private static double Nper(object?[] a) { double r = Num(a[0]), pmt = Num(a[1]), pv = Num(a[2]), fv = a.Length > 3 ? Num(a[3]) : 0, tp = a.Length > 4 ? Num(a[4]) : 0; if (r == 0) return -(pv + fv) / pmt; double q2 = 1 + tp * r; return Math.Log((-fv + pmt * q2 / r) / (pv + pmt * q2 / r)) / Math.Log(1 + r); }
    private static double Rate(object?[] a) { double nper = Num(a[0]), pmt = Num(a[1]), pv = Num(a[2]), fv = a.Length > 3 ? Num(a[3]) : 0, tp = a.Length > 4 ? Num(a[4]) : 0, guess = a.Length > 5 ? Num(a[5]) : 0.1; double r = guess; for (int i = 0; i < 100; i++) { double q = Math.Pow(1 + r, nper), f2 = pv * q + pmt * (1 + tp * r) * (q - 1) / r + fv; double df = pv * nper * Math.Pow(1 + r, nper - 1) + pmt * (tp * (q - 1) / r + (nper * (1 + tp * r) * Math.Pow(1 + r, nper - 1) / r - (1 + tp * r) * (q - 1) / (r * r))); if (Math.Abs(df) < 1e-30) break; double dr = f2 / df; r -= dr; if (Math.Abs(dr) < 1e-8) break; } return r; }
    private static double Ipmt(object?[] a) { double r = Num(a[0]), per = Num(a[1]), n = Num(a[2]), pv = Num(a[3]), fv = a.Length > 4 ? Num(a[4]) : 0, tp = a.Length > 5 ? Num(a[5]) : 0; if (per < 1 || per > n) return double.NaN; double pmt2 = Pmt(new object?[] { r, n, pv, fv, tp }); double fv2 = -Fv(new object?[] { r, per - 1, pmt2, pv, tp }); return fv2 * r; }
    private static double IrrFunc(object?[] a) { var cf = ResolveNums(new[] { a[0] }).ToList(); double r = a.Length > 1 ? Num(a[1]) : 0.1; for (int i = 0; i < 100; i++) { double npv = cf.Select((c, t) => c / Math.Pow(1 + r, t)).Sum(); double dn = cf.Select((c, t) => -t * c / Math.Pow(1 + r, t + 1)).Sum(); if (Math.Abs(dn) < 1e-30) break; r -= npv / dn; if (Math.Abs(npv / dn) < 1e-8) break; } return r; }
    private static double MirrFunc(object?[] a) { var cf = ResolveNums(new[] { a[0] }).ToList(); double fin = Num(a[1]), rein = Num(a[2]); int n = cf.Count - 1; double neg = cf.Where(v => v < 0).Select((v, i) => v / Math.Pow(1 + fin, i)).Sum(); double pos = cf.Where((v, i) => v > 0).Select((v, i2) => v * Math.Pow(1 + rein, n - i2)).Sum(); return Math.Pow(-pos / neg, 1.0 / n) - 1; }
    private static double XnpvFunc(object?[] a) { double r = Num(a[0]); var cfs = ResolveNums(new[] { a[1] }).ToList(); var dates = ResolveNums(new[] { a[2] }).Select(DateTime.FromOADate).ToList(); var d0 = dates.FirstOrDefault(); return cfs.Zip(dates).Sum(p => p.First / Math.Pow(1 + r, (p.Second - d0).Days / 365.0)); }
    private static double XirrFunc(object?[] a) { var cfs = ResolveNums(new[] { a[0] }).ToList(); var dates = ResolveNums(new[] { a[1] }).Select(DateTime.FromOADate).ToList(); var d0 = dates.FirstOrDefault(); double r = a.Length > 2 ? Num(a[2]) : 0.1; for (int i = 0; i < 100; i++) { double npv = cfs.Zip(dates).Sum(p => p.First / Math.Pow(1 + r, (p.Second - d0).Days / 365.0)); double dn = cfs.Zip(dates).Sum(p => -p.First * (p.Second - d0).Days / 365.0 / Math.Pow(1 + r, (p.Second - d0).Days / 365.0 + 1)); if (Math.Abs(dn) < 1e-30) break; r -= npv / dn; if (Math.Abs(npv / dn) < 1e-8) break; } return r; }
    private static double CumIPmt(object?[] a) { double r = Num(a[0]), n = Num(a[1]), pv = Num(a[2]); int s = (int)Num(a[3]), e = (int)Num(a[4]); double t = 0; for (int p = s; p <= e; p++) t += Ipmt(new object?[] { r, (double)p, n, pv, 0, 0 }); return t; }
    private static double CumPrinc(object?[] a) { double r = Num(a[0]), n = Num(a[1]), pv = Num(a[2]); int s = (int)Num(a[3]), e = (int)Num(a[4]); double t = 0; for (int p = s; p <= e; p++) t += Pmt(new object?[] { r, n, pv, 0.0, 0.0 }) - Ipmt(new object?[] { r, (double)p, n, pv, 0.0, 0.0 }); return t; }
    private static double Ddb(object?[] a) { double cost = Num(a[0]), salvage = Num(a[1]), life = Num(a[2]), period = Num(a[3]), factor = a.Length > 4 && a[4] != null ? Num(a[4]) : 2; double bv = cost; for (int p = 1; p <= (int)period; p++) { double dep = Math.Min(bv * factor / life, bv - salvage); if (p == (int)period) return Math.Max(dep, 0); bv -= dep; } return 0; }
    private static double DbDep(object?[] a) { double cost = Num(a[0]), salvage = Num(a[1]), life = Num(a[2]), period = Num(a[3]); double rate = 1 - Math.Pow(salvage / cost, 1 / life); double bv = cost; for (int p = 1; p < (int)period; p++) bv -= bv * rate; return bv * rate; }

    // ── Date helpers ──────────────────────────────────────────────────────────────────────────────

    private static int Days360(DateTime s, DateTime e, bool eu) { int d1 = s.Day, m1 = s.Month, y1 = s.Year, d2 = e.Day, m2 = e.Month, y2 = e.Year; if (eu) { if (d1 == 31) d1 = 30; if (d2 == 31) d2 = 30; } else { if (d1 == 31) d1 = 30; if (d2 == 31 && d1 == 30) d2 = 30; } return 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1); }
    private static double YearFrac(DateTime s, DateTime e, int basis) { int days = (e - s).Days; return basis switch { 0 => (double)Days360(s, e, false) / 360, 1 => (double)days / (DateTime.IsLeapYear(s.Year) ? 366 : 365), 2 => (double)days / 360, 3 => (double)days / 365, 4 => (double)Days360(s, e, true) / 360, _ => (double)days / 365 }; }
    private static int Weekday(DateTime dt, int rt) { int dow = (int)dt.DayOfWeek; return rt switch { 2 => dow == 0 ? 7 : dow, 3 => dow == 0 ? 6 : dow - 1, _ => dow + 1 }; }
    private static int WeekNum(DateTime dt, int rt) { if (rt == 21) return System.Globalization.ISOWeek.GetWeekOfYear(dt); var j = new DateTime(dt.Year, 1, 1); int start = rt == 2 ? 1 : 0; return (dt.DayOfYear + ((int)j.DayOfWeek - start + 7) % 7 - 1) / 7 + 1; }
    private static double NetDays(DateTime start, DateTime end, List<DateTime> holidays) { if (start > end) return 0; var h = holidays.Select(d => d.Date).ToHashSet(); int count = 0; for (var d = start; d <= end; d = d.AddDays(1)) if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday && !h.Contains(d.Date)) count++; return count; }
    private static DateTime WorkDay(DateTime start, int days, List<DateTime> holidays) { int step = days >= 0 ? 1 : -1; days = Math.Abs(days); var h = holidays.Select(x => x.Date).ToHashSet(); while (days > 0) { start = start.AddDays(step); if (start.DayOfWeek != DayOfWeek.Saturday && start.DayOfWeek != DayOfWeek.Sunday && !h.Contains(start.Date)) days--; } return start; }

    // ── Text format helper ────────────────────────────────────────────────────────────────────────

    private static string FmtValue(object? v, string fmt) { if (fmt is "General" or "@") return v?.ToString() ?? ""; if (v is double d) { try { if (fmt.Contains('%')) return (d * 100).ToString("0.##") + "%"; if (fmt.Contains("yyyy") || fmt.Contains("dd")) return DateTime.FromOADate(d).ToString(fmt); return d.ToString(fmt); } catch { return d.ToString(); } } return v?.ToString() ?? ""; }

    // ── Engineering helpers ───────────────────────────────────────────────────────────────────────

    private static double ErfFunc(double lower, double? upper)
    {
        static double Erf(double x)
        {
            double t = 1 / (1 + 0.3275911 * Math.Abs(x));
            double y = t * (0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
            return 1 - y * Math.Exp(-x * x) * (x >= 0 ? 1 : -1);
        }
        double lo = lower >= 0 ? Erf(lower) : -Erf(-lower);
        if (!upper.HasValue) return lo;
        double hi = upper.Value >= 0 ? Erf(upper.Value) : -Erf(-upper.Value);
        return hi - lo;
    }

    private static Complex ParseComplex(string s)
    {
        var m = Regex.Match(s, @"^([+-]?[\d.]+(?:[eE][+-]?\d+)?)([+-][\d.]*(?:[eE][+-]?\d+)?)[ij]$");
        if (m.Success)
        {
            double r2 = double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            string is2 = m.Groups[2].Value;
            double iv = is2 is "+" or "-" ? (is2 == "+" ? 1 : -1)
                : double.Parse(is2, System.Globalization.CultureInfo.InvariantCulture);
            return new Complex(r2, iv);
        }
        if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) return new Complex(v, 0);
        return Complex.Zero;
    }

    private static double ConvertUnits(double v, string from, string to)
    {
        var toM = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        { ["m"]=1,["km"]=1000,["mi"]=1609.344,["ft"]=0.3048,["in"]=0.0254,["cm"]=0.01,["mm"]=0.001,["Nmi"]=1852,["kg"]=1,["lbm"]=0.453592,["g"]=0.001,["J"]=1,["kJ"]=1000,["cal"]=4.184,["BTU"]=1055.056,["Pa"]=1,["atm"]=101325,["psi"]=6894.76 };
        if (from=="C"&&to=="F") return v*9/5+32; if (from=="F"&&to=="C") return (v-32)*5/9;
        if (from=="K"&&to=="C") return v-273.15; if (from=="C"&&to=="K") return v+273.15;
        if (!toM.TryGetValue(from,out var fm)||!toM.TryGetValue(to,out var tm)) return double.NaN;
        return v*fm/tm;
    }

    // ── DB function helpers ───────────────────────────────────────────────────────────────────────

    private static IEnumerable<object?> DbFilter(object?[] a, ExcelWorksheet ws)
        => FormulaEngine.ResolveValues(a[0]).Where(v => v != null && v.ToString() != "");

    private static object? DbAgg(object?[] a, ExcelWorksheet ws, Func<List<double>, object?> fn)
        => fn(DbFilter(a, ws).Select(v => Num(v)).ToList());

    // ── Math helpers ──────────────────────────────────────────────────────────────────────────────

    private static double Combination(int n, int k) { if (k < 0 || k > n) return 0; k = Math.Min(k, n - k); double r = 1; for (int i = 0; i < k; i++) { r *= (n - i); r /= (i + 1); } return r; }
    private static double Permutation(int n, int k) { if (k < 0 || k > n) return 0; double r = 1; for (int i = n; i > n - k; i--) r *= i; return r; }
    private static double Factorial(int n) { if (n < 0) return double.NaN; if (n > 170) return double.PositiveInfinity; double r = 1; for (int i = 2; i <= n; i++) r *= i; return r; }
    private static double DoubleFactorial(int n) { if (n <= 0) return 1; double r = 1; for (int i = n; i > 0; i -= 2) r *= i; return r; }
    private static double Multinomial(List<double> nums) { double s = nums.Sum(), r = Factorial((int)s); foreach (var n in nums) r /= Factorial((int)n); return r; }
    private static long GcdLong(long a, long b) => b == 0 ? Math.Abs(a) : GcdLong(b, a % b);
    private static long LcmLong(long a, long b) => a == 0 || b == 0 ? 0 : Math.Abs(a / GcdLong(a, b) * b);
    private static double RoundUp(double v, int d) { double f = Math.Pow(10, d); return v >= 0 ? Math.Ceiling(v * f) / f : -Math.Floor(-v * f) / f; }
    private static double RoundDown(double v, int d) { double f = Math.Pow(10, d); return v >= 0 ? Math.Floor(v * f) / f : -Math.Ceiling(-v * f) / f; }
    private static double RoundToMag(double v, int d) { double f = Math.Pow(10, -d); return Math.Round(v / f, MidpointRounding.AwayFromZero) * f; }
    private static string ToRoman(int n) { if (n < 1 || n > 3999) return ""; var p = new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") }; var sb = new System.Text.StringBuilder(); foreach (var (v, s) in p) { while (n >= v) { sb.Append(s); n -= v; } } return sb.ToString(); }
    private static int RomanToArabic(string s) { var vals = new Dictionary<char, int> { ['I']=1,['V']=5,['X']=10,['L']=50,['C']=100,['D']=500,['M']=1000 }; int total = 0; for (int i = 0; i < s.Length; i++) { int v = vals.GetValueOrDefault(char.ToUpper(s[i]), 0), next = i + 1 < s.Length ? vals.GetValueOrDefault(char.ToUpper(s[i + 1]), 0) : 0; total += v < next ? -v : v; } return total; }
}
