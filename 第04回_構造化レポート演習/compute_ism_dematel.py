#!/usr/bin/env python3
"""ISM・DEMATEL 計算（2026_ISM-DEMATEL.xlsm と同等ロジック）"""
import json
import numpy as np

ITEMS = [
    "GPA・単位取得",
    "試験・レポートの質",
    "学習時間の確保",
    "睡眠・体調管理",
    "授業内容の理解",
    "学習意欲・ストレス",
    "バイト・課外の負荷",
    "学習環境（集中）",
]

# 隣接行列 A[i,j]=1 : 要素 i が j に影響（行→列）
A = np.array([
    [0,0,0,0,0,0,0,0],  # 1 GPA
    [0,0,0,0,0,0,0,0],  # 2 試験質
    [0,1,0,0,1,0,0,0],  # 3 学習時間 →2,5
    [0,1,0,0,1,1,0,0],  # 4 睡眠 →2,5,6
    [0,1,0,0,0,0,0,0],  # 5 理解 →2
    [0,0,1,0,1,0,0,0],  # 6 意欲 →3,5
    [0,0,1,1,0,0,0,0],  # 7 バイト →3,4
    [0,0,1,0,1,0,0,0],  # 8 環境 →3,5
], dtype=int)

# 試験・レポートの質 → GPA
A[1, 0] = 1

# DEMATEL 直接影響行列（0〜4）
X = np.array([
    [0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0],
    [0,4,0,0,2,0,0,0],  # 3→2,5
    [0,3,0,0,2,2,0,0],  # 4→2,5,6
    [0,4,0,0,0,0,0,0],  # 5→2
    [0,0,3,0,2,0,0,0],  # 6→3,5
    [0,0,4,3,0,0,0,0],  # 7→3,4
    [0,0,2,0,2,0,0,0],  # 8→3,5
], dtype=float)
X[1, 0] = 4  # 2→1


def bool_mult(M, B):
    n = M.shape[0]
    C = np.zeros((n, n), dtype=int)
    for i in range(n):
        for j in range(n):
            C[i, j] = 1 if np.dot(M[i], B[:, j]) > 0 else 0
    return C


def ism(A):
    n = A.shape[0]
    ApI = A.copy()
    np.fill_diagonal(ApI, 1)
    B = ApI.copy()
    while True:
        C = bool_mult(B, ApI)
        if np.array_equal(C, B):
            break
        B = C
    reach = B

    hierarchy = []
    Rc = reach.copy()
    remaining = n
    while remaining > 0:
        level = []
        for i in range(n):
            if not np.any(Rc[i]):
                continue
            # R(i) subset of A(i): all j with reach[i,j]=1 also have reach[j,i]=1
            ri = set(j for j in range(n) if Rc[i, j])
            ai = set(j for j in range(n) if Rc[j, i])
            if ri <= ai:
                level.append(i + 1)
                remaining -= 1
        for idx in level:
            Rc[idx - 1, :] = 0
            Rc[:, idx - 1] = 0
        if level:
            hierarchy.append(level)
        else:
            break

    setR = [[j + 1 for j in range(n) if reach[i, j]] for i in range(n)]
    setA = [[j + 1 for j in range(n) if reach[j, i]] for i in range(n)]
    return ApI, reach, setR, setA, hierarchy


def skeleton(A):
    """推移的冗長辺を除いた骨格行列（別経路があれば i→j を削除）"""
    n = A.shape[0]
    ApI = A.copy()
    np.fill_diagonal(ApI, 1)
    reach = ApI.copy()
    while True:
        C = bool_mult(reach, ApI)
        if np.array_equal(C, reach):
            break
        reach = C
    sk = A.copy()
    for i in range(n):
        for j in range(n):
            if not A[i, j]:
                continue
            for k in range(n):
                if k == i or k == j:
                    continue
                if reach[i, k] and reach[k, j]:
                    sk[i, j] = 0
                    break
    return sk


def dematel(X):
    n = X.shape[0]
    row_sums = X.sum(axis=1)
    col_sums = X.sum(axis=0)
    mx = max(row_sums.max(), col_sums.max())
    N = X / mx
    I = np.eye(n)
    T = N @ np.linalg.inv(I - N)
    D = T.sum(axis=1)
    R = T.sum(axis=0)
    return N, T, D, R, D + R, D - R


def main():
    ApI, reach, setR, setA, hierarchy = ism(A)
    sk = skeleton(A)
    N, T, D, R, DpR, DmR = dematel(X)

    out = {
        "items": ITEMS,
        "A": A.tolist(),
        "ApI": ApI.tolist(),
        "reach": reach.tolist(),
        "setR": setR,
        "setA": setA,
        "hierarchy": hierarchy,
        "skeleton": sk.tolist(),
        "X": X.tolist(),
        "N": N.round(4).tolist(),
        "T": T.round(4).tolist(),
        "D": D.round(4).tolist(),
        "R": R.round(4).tolist(),
        "DpR": DpR.round(4).tolist(),
        "DmR": DmR.round(4).tolist(),
    }
    path = "/Users/shuta/psi/s1/⭐️0526火3 金3sys-kogakukiso工学基礎 5/0516/results.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("=== 階層 ===")
    for lv, elems in enumerate(hierarchy):
        names = [f"{e}:{ITEMS[e-1]}" for e in elems]
        print(f"Level {lv+1}: {', '.join(names)}")
    print("\n=== DEMATEL D-R (影響度) 降順 ===")
    order = sorted(range(len(ITEMS)), key=lambda i: -DmR[i])
    for i in order:
        role = "原因系" if DmR[i] > 0.01 else ("結果系" if DmR[i] < -0.01 else "媒介")
        print(f"  {i+1} {ITEMS[i]}: D-R={DmR[i]:.3f} D+R={DpR[i]:.3f} [{role}]")
    print("\n骨格行列の辺:")
    for i in range(len(ITEMS)):
        for j in range(len(ITEMS)):
            if sk[i, j]:
                print(f"  {i+1}→{j+1}: {ITEMS[i]} → {ITEMS[j]}")


if __name__ == "__main__":
    main()
