function fillArray(length = 10, shift = 2, fill = 1){
  
  let a = new Array(length);
  a.fill(fill);
  let count = shift;
  for(i in a)
  {
    if(count==6){a[i]=""}
    if(count==7){a[i]="";count=1;count--;}
    count++;
  }

  return a;  
}
