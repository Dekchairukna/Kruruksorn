# Railway build fix: mise Python GitHub attestation

ถ้า Railway build ล้มที่ข้อความประมาณนี้:

```text
mise ERROR Failed to install core:python@3.11.9: No GitHub artifact attestations found
To disable attestation verification, set MISE_PYTHON_GITHUB_ATTESTATIONS=false
or add python.github_attestations = false under [settings] in mise.toml
```

ให้ทำอย่างใดอย่างหนึ่ง:

1. กดปุ่ม **Disable attestation check** บนหน้า Railway แล้ว Redeploy
2. หรือเพิ่ม Railway Variable:

```env
MISE_PYTHON_GITHUB_ATTESTATIONS=false
```

ในชุดไฟล์นี้เพิ่ม `mise.toml` และ `.mise.toml` ให้แล้ว:

```toml
[settings]
python.github_attestations = false
```

หลังอัปขึ้น GitHub แล้วให้ Redeploy ใหม่
