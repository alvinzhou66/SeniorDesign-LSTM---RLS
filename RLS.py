import numpy as np
import matplotlib.pylab as plt
import padasip as pa
import pandas as pd

dataset_train = pd.read_csv(r'C:\Users\dell\Desktop\Xian.csv')
input_set = dataset_train.iloc[:, 1:2].values

#cutting dataset into certain chunk-size,n is 2,3,5,7,10
X = []
y_real= []
n = 2
for i in range(n, len(input_set)):
 X.append(input_set[i-n:i, 0])
 y_real.append(input_set[i, 0])
X, y_real = np.array(X), np.array(y_real)

# identification
f = pa.filters.FilterRLS(n, mu=0.98)
y, e, w = f.run(y_real,X)

# show results

squareerror = []
for val in e:
    squareerror.append(val * val)
mse = sum(squareerror) / len(squareerror)
cor = np.corrcoef(np.array(y_real), np.array(y))
spm = pd.DataFrame({'real': y_real, 'prediction': y}).corr('spearman')
print(mse)
print(cor)
print(spm)


plt.figure(figsize=(15,9))
plt.title("Xian");plt.xlabel("days")
plt.plot(y_real,"b", label="real")
plt.plot(y,"r", label="prediction");plt.legend()
plt.tight_layout()
plt.show()

