公式說明
===
水平力分配公式
===
由`鐵路橋梁耐震數計規範2.7節`總橫力分配方法為:  
$$
\begin{aligned}
p_e(x) &= \frac{\Sigma w(x)v_s(x)}{\Sigma w(x) v{^2}{_s}(x)}w(x)v_s(x)\frac{V}{W+L_E} \\
\end{aligned}
$$  
因此總分配力為:
$$
\begin{aligned}
\Sigma p_e(x) &= \frac{\Sigma w(x)v_s(x)}{\Sigma w(x) v{^2}{_s}(x)}\Sigma w(x)v_s(x)\frac{V}{W+L_E} 
\end{aligned}
$$ 
$\frac{V}{W+L_E}$前的項目可寫成:
$$
R = \frac{[\Sigma w(x)v_s(x)]^2}{\Sigma w(x) v{^2}{_s}(x)}
$$
若由柯西不等式知:  
$$
\begin{aligned}
Define \quad \hat{u} &= \frac{\Sigma wu}{\Sigma w} \quad (為加權平均概念)\\
E[u^2]&= 以u為權之二次平均
\end{aligned}
$$
所以原式  
分子為$(\Sigma wu)^2$，分母為$\Sigma wu^2 = \Sigma w E[u^2]$
故
$$
\begin{aligned}
R &= \frac{[\Sigma w(x)v_s(x)]^2}{\Sigma w(x) v{^2}{_s}(x)} \\
&=(\Sigma w)\frac{\hat{u^2}}{E[u^2]}
\end{aligned}
$$
但
$$
E[u^2] \geq \hat{u^2}
$$
所以
$$
\begin{aligned}
&0 \leq \frac{[\Sigma w(x)v_s(x)]^2}{\Sigma w(x) v{^2}{_s}(x)} \leq (W+L_E) \\
\Rightarrow & 0 \leq R \leq (W+L_E) \\
\end{aligned}
$$
故得證
$$
\begin{aligned}
\Sigma p_e(x) &= \frac{\Sigma w(x)v_s(x)}{\Sigma w(x) v{^2}{_s}(x)}\Sigma w(x)v_s(x)\frac{V}{W+L_E} \\
&= R\frac{V}{W+L_E} \\
& \neq V
\end{aligned}
$$

振態疊加
===
因公式為振態疊加，故僅以第一振態疊加幾乎不是100%V，因此計算得$\Sigma p_e$會低於$V$。
已以SAP雙自由度模型證明，依規範推出結果近似第一振態計算結果，但$\Sigma p_e$等於$V$要兩個模態疊加。
