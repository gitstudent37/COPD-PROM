# [file name]: Heatmap.py
# [file content begin]
# -*- coding:utf-8 -*-
"""
Copyright (C) 2025 SXMU

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap


class HeatmapPainter:
    def __init__(self, corr_matrix, p_matrix, significance_levels=[0.05, 0.01, 0.001]):
        """
        Correlation heatmap with significance markers
        :param corr_matrix: Correlation matrix（DataFrame or numpy array）
        :param p_matrix: p-value matrix（DataFrame）
        :param significance_levels: Significance level, defaults [0.05, 0.01, 0.001].
        :return:
        """
        self.corr_matrix = corr_matrix
        self.p_matrix = p_matrix
        self.significance_levels = significance_levels
        self._param_check()
        self._create_annotations()

    def _param_check(self):
        # Confirming the type of input parameters
        if not isinstance(self.corr_matrix, pd.DataFrame):
            if isinstance(self.corr_matrix, np.ndarray):
                self.corr_matrix = pd.DataFrame(self.corr_matrix)
            else:
                raise ValueError('corr_matrix must be DataFrame or ndarray')
        if not isinstance(self.p_matrix, pd.DataFrame):
            if isinstance(self.p_matrix, np.ndarray):
                self.p_matrix = pd.DataFrame(self.p_matrix)
            else:
                raise ValueError('p_matrix must be DataFrame or ndarray')

        # Checking the input parameters
        if len(self.significance_levels) != 3:
            raise ValueError('Significance levels of three floating-point thresholds are required')
        for item in self.significance_levels:
            if not isinstance(item, float):
                raise ValueError('Significance levels of floating-point thresholds are required')

        # Calculating the actual comparison thresholds and regenerate the set of comparison thresholds
        num_compared = self.p_matrix.shape[0] * self.p_matrix.shape[1]
        alpha = self.significance_levels[0]
        alpha_mod = alpha / num_compared
        self.significance_levels.append(alpha_mod)
        self.significance_levels.sort(reverse=True)
        del_end_ind = self.significance_levels.index(alpha_mod)
        del self.significance_levels[: del_end_ind]

        print(
            f'The total number of comparisons is {num_compared}',
            f'Bonferroni method: alpha before and after correction are {alpha} and {alpha_mod}',
            f'The final determined floating-point threshold is: {self.significance_levels}',
            sep='\n'
        )

    def _create_annotations(self):
        annotations = pd.DataFrame('',
                                   index=self.corr_matrix.index,
                                   columns=self.corr_matrix.columns,
                                   dtype=str)
        for i in range(self.p_matrix.shape[0]):
            for j in range(self.p_matrix.shape[1]):
                p = self.p_matrix.iloc[i, j]
                if len(self.significance_levels) == 3:
                    if p < self.significance_levels[2]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f') + '***'
                    elif p < self.significance_levels[1]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f') + '**'
                    elif p < self.significance_levels[0]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f') + '*'
                    else:
                        annotations.iloc[i, j] = ''
                elif len(self.significance_levels) == 2:
                    if p < self.significance_levels[1]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f') + '**'
                    elif p < self.significance_levels[0]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f') + '*'
                    else:
                        annotations.iloc[i, j] = ''
                else:
                    if p < self.significance_levels[0]:
                        annotations.iloc[i, j] = format(self.corr_matrix.iloc[i, j], '.3f')
                    else:
                        annotations.iloc[i, j] = ''
        self.annotations = annotations

    def plot_correlation_heatmap(self):
        # Set canvas size
        plt.figure(figsize=(1.1 * self.corr_matrix.shape[1], 1.1 * self.corr_matrix.shape[0]))
        plt.subplots_adjust(left=0.11, right=1.0, bottom=0.2, top=1.0)

        # Set the heatmap color bands
        custom_cmap = LinearSegmentedColormap.from_list(
            "custom_cmap",
            ["#104e8b", "#5f89b1", "#ffffff", "#e5b5b5", "#b22222"],
            N=1024
        )

        # Drawing a heatmap
        ax = sns.heatmap(
            self.corr_matrix,
            annot=self.annotations,
            fmt='',
            # cmap='coolwarm',
            cmap=custom_cmap,
            center=0,
            cbar_kws={
                'label': 'Correlation Coefficient',
                'shrink': 1,
                'aspect': 15,
                'pad': 0.02,
            },  # Colorbar related properties
            annot_kws={
                "size": 13,
            },  # Annotation-related properties
            linewidths=0.5,
            linecolor='white',
        )

        # Font size of the colorbar label
        colorbar = ax.collections[0].colorbar
        colorbar.set_label('Correlation Coefficient', fontsize=13)  # Label and font size

        # Axis font
        plt.xticks(
            rotation=30, ha='right',
            fontsize=13,
            # fontweight='bold',
        )
        plt.yticks(
            fontsize=13,
            # fontweight='bold',
        )

        # plt.title('Correlation Heatmap with Significance')
        plt.show()


if __name__ == '__main__':
    df_corr = pd.read_csv('data_heatmap_demo/cor_matrix_demo.csv', index_col=0).T
    df_p = pd.read_csv('data_heatmap_demo/p_matrix_demo.csv', index_col=0).T
    hp = HeatmapPainter(df_corr, df_p)
    hp.plot_correlation_heatmap()

# [file content end]