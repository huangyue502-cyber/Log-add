# Log-add
Matching data through stratigraphic depth constraints
import os
import pandas as pd

# 定义文件夹路径和输出文件路径
folder_a = 'D:\\python\\projects\\pythonProject1\\excel test\\12'
folder_b = 'D:\\python\\projects\\pythonProject1\\excel test\\11'
output_file = 'D:\\python\\projects\\pythonProject1\\excel test\\constrained_averages.xlsx'

# 初始化结果 DataFrame
result_df = pd.DataFrame(columns=["井号", "X一段平均值", "X二段平均值", "X三段平均值", "XXXX平均值"])

# 遍历文件夹 A 中的 Excel 文件
for file_a_name in os.listdir(folder_a):
    if file_a_name.endswith(".xlsx"):
        file_a_path = os.path.join(folder_a, file_a_name)

        try:
            # 读取文件夹 A 中的 Excel 文件
            df_a = pd.read_excel(file_a_path)

            # 遍历文件夹 A 中的每一行数据（即每一个井号）
            for _, row in df_a.iterrows():
                well_id = str(row[0])  # 获取井号

                # 定义深度范围
                depth_constraints = {
                    "X一段平均值": (row[2], row[1]),  # 第二列为底深，第三列为顶深
                    "X二段平均值": (row[4], row[3]),  # 第四列为底深，第五列为顶深
                    "X三段平均值": (row[6], row[5]),  # 第六列为底深，第七列为顶深
                    "XXXX平均值": (row[8], row[7])  # 第八列为底深，第九列为顶深
                }

                # 构造文件夹 B 中的匹配文件名
                file_b_name = f"{well_id}.xlsx"
                file_b_path = os.path.join(folder_b, file_b_name)

                # 检查文件夹 B 中是否存在匹配的文件
                if os.path.exists(file_b_path):
                    try:
                        # 读取文件夹 B 中的 Excel 文件
                        df_b = pd.read_excel(file_b_path)

                        # 初始化字典以存储结果
                        average_values = {"井号": well_id}

                        # 遍历每个深度范围，并计算深度及扩展列的平均值
                        for avg_name, (depth_min, depth_max) in depth_constraints.items():
                            # 筛选符合深度范围的行，假设深度列名称为 "Depth"
                            constrained_data = df_b[(df_b["Depth"] >= depth_min) & (df_b["Depth"] <= depth_max)]

                            # 计算每列的平均值
                            if not constrained_data.empty:
                                avg_values = constrained_data.mean()

                                # 为每个列的平均值创建命名，存入 average_values
                                for col_name, avg_value in avg_values.items():
                                    avg_col_name = f"{avg_name}_{col_name}平均值"
                                    average_values[avg_col_name] = avg_value
                            else:
                                # 如果没有符合条件的行，则填入 None
                                for col_name in df_b.columns:
                                    if col_name != "Depth":
                                        avg_col_name = f"{avg_name}_{col_name}平均值"
                                        average_values[avg_col_name] = None

                        # 将结果追加到结果 DataFrame
                        result_df = pd.concat([result_df, pd.DataFrame([average_values])], ignore_index=True)
                        print(f"已处理井号: {well_id}")

                    except Exception as e:
                        print(f"处理文件 {file_b_name} 时出错: {e}")

        except Exception as e:
            print(f"处理文件 {file_a_name} 时出错: {e}")

# 将结果保存到新的 Excel 文件中
result_df.to_excel(output_file, index=False)
print(f"所有数据已提取并保存到 {output_file}")






