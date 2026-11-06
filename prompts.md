这次已经非常接近首次成功迁移了。AALC2201 的完整 Audit 已经 PASS，现在唯一确认的阻塞原因是缺少 `ASAP_INTERNAL_LABEL_ID` 环境变量。

不需要重新修改 Formula Remapping 逻辑，也不需要重新优化 Dependency Scan。

## 一、当前执行状态

| 项目                      | 状态              |
| ----------------------- | --------------- |
| Dependency Audit        | &#x20;PASS      |
| 48,300 个依赖识别            | &#x20;完成        |
| `Comparison!AC4` 公式支持   | &#x20;已修复       |
| `COUNTA` 与 Defined Name | &#x20;已处理       |
| 自动化测试                   | &#x20;33 Passed |
| Internal Label API      | &#x20;可正常访问     |
| Internal Label 环境变量     | &#x20;未配置       |
| 最终 XLSB 迁移              | 尚未完成            |

## 二、直接解决：设置已知的 Internal Label GUID

你此前已经确认过公司的 Internal Label ID：

`dd181445-6ec4-4473-9810-00785f082df0`

在当前 VS Code PowerShell Terminal 中执行：

```

$env:ASAP_INTERNAL_LABEL_ID = "dd181445-6ec4-4473-9810-00785f082df0"

Write-Host "Internal Label ID:" $env:ASAP_INTERNAL_LABEL_ID

```

随后运行：

```

python tools/remap_proof.py --deal AALC2201 --com-timeout 1800

```

这里假设你在 Repository 根目录执行，且当前 Copilot 代码仍沿用 `tools/remap_proof.py` 路径。

如果你希望以后新开的 PowerShell Terminal 也能识别这个环境变量，可以使用：

```
[Environment]::SetEnvironmentVariable(
    "ASAP_INTERNAL_LABEL_ID",
    "dd181445-6ec4-4473-9810-00785f082df0",
    "User"
)
```

持久化设置只对之后新开的进程生效；当前 Terminal 仍需使用第一段 `$env:` 设置。

## 三、我建议对工具做一个永久性的小改进

现在你已经确认：

- 原始 Proof 自带 `Internal` Label。
- `GetLabel()` 可以正常返回。
- 实际 LabelId 与预期一致。

那么在迁移时，程序应当优先保留并验证原文件已有的 Internal Label，而不是不论情况如何都要求重新配置环境变量。

建议给 Copilot 这段简短 Prompt：

Update the existing Sensitivity Label handling in `remap_proof.py`.

The AALC2201 audit is now PASS, but migration is blocked because `ASAP_INTERNAL_LABEL_ID` is not configured.

The original workbook already has the correct Internal sensitivity label, previously verified as:

`dd181445-6ec4-4473-9810-00785f082df0`

Implement the following minimal change:

1. Inspect the original workbook's sensitivity label through `GetLabel()`.
2. If the existing LabelId matches the approved Internal LabelId, preserve it without unnecessary reassignment.
3. Do not require label-assignment environment variables when no label assignment is necessary.
4. If the workbook is unlabeled, require authorized configuration before applying Internal.
5. If the workbook has a different label, fail closed rather than overwrite or downgrade it.
6. After saving and reopening the XLSB, verify that the persisted LabelId still matches Internal.
7. Keep the approved expected LabelId in an appropriate local configuration, without unnecessarily publishing tenant-specific configuration in the public repository.
8. Preserve all other migration functionality and existing safeguards.

Do not modify formula mapping, dependency scanning, versioning, financial calculations, or workbook-preservation logic.

Add focused regression tests, then execute the actual AALC2201 migration using the existing PASS audit evidence where it is still valid.

Report the final filename, formula migration count, financial comparison result, workbook preservation result, and persisted sensitivity label.

Do not claim success until the final versioned XLSB has been saved, reopened, and fully validated.

## 四、这轮最应该关注什么？

下一次成功执行后，应该看到类似以下结果：

AALC2201 Migration

待验证

Dependency Audit

PASS

Formula Remapping

Pending

Financial Tie-out

Pending

XLSB Save / Reopen

Pending

Sensitivity Label

Internal — Pending

Version Increment

Pending

我的建议：先直接设置环境变量并尝试完整迁移。 这是当前最短的解决路径。上述永久性改进可以随后实施，以免你每次处理不同 Deal 都要重复配置同一个 Internal Label。
