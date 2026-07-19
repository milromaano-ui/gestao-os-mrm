# Gestão de OS — MRM Personal Car

App para cadastrar carros, peças, serviços e gerar a OS do cliente,
com painel de custos/lucro. Os dados agora ficam salvos de verdade
no Supabase (antes ficavam só na memória do Claude).

## 1. Configurar o Supabase

1. Entre no seu projeto em https://supabase.com/dashboard
2. Vá em **SQL Editor** → **New query**
3. Cole o conteúdo do arquivo `supabase/schema.sql` deste projeto e clique em **Run**
   (isso cria a tabela `kv_store`, onde o app guarda as OS, os clientes e os outros custos)
4. Vá em **Project Settings > API**
5. Copie:
   - **Project URL** → vai virar `NEXT_PUBLIC_SUPABASE_URL`
   - **anon public key** → vai virar `NEXT_PUBLIC_SUPABASE_ANON_KEY`

## 2. Rodar localmente (opcional, mas recomendado antes de publicar)

Precisa ter o [Node.js](https://nodejs.org) instalado no computador.

```bash
cd gestao-os-mrm
npm install
cp .env.local.example .env.local
```

Abra o `.env.local` e cole a URL e a anon key do Supabase. Depois:

```bash
npm run dev
```

Abra http://localhost:3000 e teste: cadastre uma OS, feche o navegador,
abra de novo — ela precisa continuar lá.

## 3. Publicar no GitHub

1. Crie um repositório novo (pode ser privado) no GitHub
2. Dentro da pasta `gestao-os-mrm`:

```bash
git init
git add .
git commit -m "Primeira versão do app de gestão de OS"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/gestao-os-mrm.git
git push -u origin main
```

## 4. Publicar na Vercel

1. Entre em https://vercel.com/new
2. Importe o repositório `gestao-os-mrm` que você acabou de subir
3. Antes de clicar em Deploy, abra **Environment Variables** e adicione:
   - `NEXT_PUBLIC_SUPABASE_URL` = (a URL do seu projeto Supabase)
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY` = (a anon key do seu projeto Supabase)
4. Clique em **Deploy**

Em ~1 minuto a Vercel te dá um link tipo `gestao-os-mrm.vercel.app`.
Esse é o endereço definitivo do app — pode salvar na tela inicial do
celular como um atalho.

## O que mudou em relação à versão de teste

- Os dados (OS, clientes, outros custos) agora são salvos no Supabase,
  não somem mais ao fechar o app.
- O botão "Texto pra enviar" (WhatsApp/Telegram) deve funcionar
  normalmente num navegador de verdade — o problema anterior era uma
  limitação do ambiente de teste dentro do Claude, não do código.
- Geração de PDF: o botão de imprimir usa a impressão do navegador —
  no celular ou computador, ao imprimir, dá pra escolher "Salvar como
  PDF" em vez de uma impressora física.

## Ainda não incluído (avise se quiser que eu adicione antes ou depois de publicar)

- Fechamento por período no Painel
- Registro de pagamentos/custos externos separado dos serviços
- Envio direto por WhatsApp/Telegram sem precisar copiar e colar
- Login/senha para proteger o acesso ao app
