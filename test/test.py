with open("in.DAT") as f:
    in_data = f.readlines()
with open("out.DAT") as f:
    out_data = f.readlines()
for i in range(0,len(in_data)):
    a = in_data[i][:-1].rstrip()
    b = out_data[i][:-1].rstrip()
    if a != b:
        print(a)
        print(b)

