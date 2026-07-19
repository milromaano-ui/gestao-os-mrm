import dynamic from "next/dynamic";
import Head from "next/head";

const GestaoOS = dynamic(() => import("../components/GestaoOS"), {
  ssr: false,
});

export default function Home() {
  return (
    <>
      <Head>
        <title>Gestão de OS — MRM Personal Car</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>
      <GestaoOS />
    </>
  );
}
