import re
from xlswriter import XlsWriter
t = ""
with open("test_data.txt", encoding="utf-8") as f:
    t += f.read()

def get_data(text):
    if text:
        text = text.replace("\r\n", "\n")
        text = text.replace("-\n", "")
        text = text.lower()
        start = re.search(r"(?im)^LP\.\s*Asortyment\b", text)
        if start:
            text = text[start.start():]
        pattern = r'(?m)^\s*\d+\s+[^\W\d_][^,]*,\s*[^,]*'
        patterns = re.findall(pattern, text)
        ans = []
        t_p = []
        for line in patterns:
            if len(line.split("\n"))>1:
                value = int(line.split("\n")[0].replace(" s", ""))
                ans.append(value)
            t_p.append(line.split("\n")[len(line.split("\n"))-1])

        final = []
        cnt = 0
        products = []
        for l in t_p:
            cnt+=1
            if cnt <= len(ans):
                products.append(l)
                tmp = [l, ans[cnt-1]]
                final.append(tmp)

        # data pattern = [["product" , ".....", "...."]
        # ["Amount", "...", "...."]]
        data = []
        for i in range(len(products)):
            data.append([products[i], ans[i]])
        return data


bot = XlsWriter("test.xlsx")
bot.create_tab(get_data(t))




