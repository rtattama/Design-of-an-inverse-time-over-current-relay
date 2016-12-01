clc;
clear all;
X=xlsread('C:\Users\Dell\Desktop\RELAY\Test Signal.xlsx','A:A');
Y=xlsread('C:\Users\Dell\Desktop\RELAY\Test Signal.xlsx','B:B');
Z=xlsread('C:\Users\Dell\Desktop\RELAY\Test Signal.xlsx','C:C');

%Initializing A,B,P values
B=0.17966;                                      
A=8.9341;
P=2.0983;
fz = 60;
tap_set = 5;
time_ds = 2;

%ANALYSIS OF FOUR LEVEL FAULT CURRENT WITH HARMONICS AND DC
a=0;
M=0;
sum=0;
rms=0;
irms=0;
c=0;
trp3=0;
tr3=0;
I=0;
m=0;
h=0;
p=0;
cf=0;
opt_time=0;

%Cosine Filter is used to filter DC Harmonics
for m=1:16                                                                
    cf(m)=cos((pi*(m-1))/8);
end

%Calculating the co efficients of the filter
for n=1:4985                                  
    p=1;
    for o=16:-1:1
        h=(Z(n+o-1)*cf(p))+h;
        p=p+1;
    end
    I(n)=h/8;
    h=0;
end

for r=1:311 
    %For the loop of one cycle = 16 samples
    for s=1:16                               
        sum=sum+((I(16*(r-1)+s))^2);         
    end
    %Sum of squares of 16 samples
    rms=sqrt(sum/16);                        
    sum=0;
    %Calculation of rms value
    irms(r)=rms;
    %Finding the multiples of tap setting
   M=rms/tap_set;                                  
   if(M>=1)
       c=c+1;
       trp3=time_ds*((A/((M^P)-1))+B)*fz; 
       %Calculation of the trip time
       tr3=tr3+(1/trp3);
       if(tr3>=1)
           %Calculation of the operating time
           opt_time=c/fz;                         
    break;
       end
   end
   
end

%TWO LEVEL FAULT CURRENT ANALYSIS
sum_1=0;
rms_1=0;
irms_1=0;
M1=0;
c_1=0;
tr_1=0;
trp_1=0;
opt_time1=0;
for a=1:311
    %For the loop of one cycle = 16 samples
    for b=1:16                               
        sum_1=sum_1+((X(16*(a-1)+b))^2);         
    end
    %Sum of squares of 16 samples
    rms_1=sqrt(sum_1/16);                        
    sum_1=0;
    %Calculation of rms value
    irms_1(a)=rms_1;
    %Finding the multiples of tap setting
   M1=rms_1/tap_set;                                  
   if(M1>=1)
       c_1=c_1+1;
       %Calculation of the trip time
       trp_1=time_ds*((A/((M1^P)-1))+B)*fz;            
       tr_1=tr_1+(1/trp_1);
       if(tr_1>=1)
           %Calculation of the operating time
           opt_time1=c_1/fz;                         
    break;
       end
   end
   
end
% FOUR LEVEL FAULT CURRENT ANALYSIS
sum_2=0;
rms_2=0;
irms_2=0;
M2=0;
c_2=0;
trp_2=0;
tr_2=0;
opt_time2=0;
for i=1:311
    %For the loop of one cycle = 16 samples
    for j=1:16                               
        sum_2=sum_2+((Y(16*(i-1)+j))^2);         
    end
    %Sum of squares of 16 samples
    rms_2=sqrt(sum_2/16);                        
    sum_2=0;
    %Calculation of rms value
    irms_2(i)=rms_2;
    %Finding the multiples of tap setting
   M2=rms_2/tap_set;                                  
   if(M2>=1)
       c_2=c_2+1;
       %Calculation of the trip time
       trp_2=time_ds*((A/((M2^P)-1))+B)*fz;            
       tr_2=tr_2+(1/trp_2);
       if(tr_2>=1)
           %Calculation of the operating time
           opt_time2=c_2/fz;                         
    break;
       end
   end
   
end
figure;
subplot(2,1,1)
plot(X);
title('TWO LEVEL CRRENT SAMPLES');
subplot(2,1,2)
plot(irms_1);
title('TWO LEVEL IRMS');

figure;
subplot(2,1,1)
plot(Y);
title('FOUR LEVEL CRRENT SAMPLES');
subplot(2,1,2)
plot(irms_2);
title('FOUR LEVEL IRMS');

figure;
subplot(3,1,1)
plot(Z);
title('FOUR LEVEL CRRENT SAMPLES WITH HARMONICS');
subplot(3,1,2)
plot(I);
title('FILTERED CURRENT');
subplot(3,1,3)
plot(irms);
title('FOUR LEVEL CRRENT IRMS WITH HARMONICS');
