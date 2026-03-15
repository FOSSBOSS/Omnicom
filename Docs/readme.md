<pre>
The protocol ENQ+device+CONTINUATION+type+datatype+value+checksum+terminator
  ENQ = ACK/NAK    
  DEVICE = FF/1:N  #BYTE
  CONTINUATION =   wait for more instructions, or dont
  TYPE = R/W/C
  Datatype: bits or registers basicly
  value=bytes vary based on datatype
  checksum=BCC 
  terminator = \0h

Alot of the docs are really old, and I misplace them. Im going  to organize them based on OEM
  
  
</pre>
