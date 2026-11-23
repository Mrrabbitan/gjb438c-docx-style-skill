# 加密资料与本地恢复

高风险原始资料已从本仓库可达历史移除或外置为环境配置。整改前所有原始分支、
标签与可达提交均完整保存在 `original-repository.bundle.age` 中。
该文件为 age 密文；本仓库仅包含公钥，解密私钥由所有者在 GitHub 之外保管。

## 恢复原始内容

使用 age v1.3.2 或兼容版本，在本仓库之外、仅自己可访问的目录操作：

```sh
umask 077
age --decrypt --identity "$AGE_IDENTITY_FILE" \
  --output /private/recovery/original.bundle \
  .protected/original-repository.bundle.age
git clone /private/recovery/original.bundle /private/recovery/original-repository
```

`/private/recovery` 为示意路径，应自行创建有权限的本地目录。
将解密后的 bundle 的 SHA256 与 `manifest.json` 的 `original_bundle_sha256` 比较。
恢复出的仓库包含原始敏感资料，只用于获授权的本地工作。不要把私钥、解密后的
bundle、原配置或原始资料重新提交到 GitHub；需要模板/数据时在独立私有目录使用。

源码保留了非敏感结构。依赖真实服务的功能必须先提供环境变量或授权配置；
原始业务样本、评估产物及项目参考模板需要先在本地解密取得。密钥的环境变量名
见对应配置与源码，缺失时不会自动恢复旧值。

文件加密不会让已公开的旧凭据失效。服务归属方仍需在服务端撤销或轮换这些凭据。
协作者应重新克隆整改后的仓库，避免把旧明文历史推回。

## 项目参考模板

原总体技术方案参考 Word 已移入密文备份。需要该模板时，先在独立私有目录恢复
原仓库，再用现有脚本的 `--template` 参数指向私有模板的绝对路径。
原默认总体方案生成路径在解密前缺少模板；整包 `validate_skill` 检查会提示这一
受保护参考文件缺失。SRS 的 P-09 模板仍保留，安装通过不等同于整包生成/审计通过。
