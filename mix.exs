defmodule Elixlsx.Mixfile do
  use Mix.Project

  def project do
    [app: :elixlsx,
     version: "0.4.0",
     elixir: "~> 1.3",
     package: package(),
     description: "a writer for XLSX spreadsheet files",
     build_embedded: Mix.env == :prod,
     start_permanent: Mix.env == :prod,
     deps: deps()]
  end

  def application do
    []
  end

  defp deps do
    [
      {:excheck, "~> 0.5", only: :test},
      {:triq, github: "triqng/triq", only: :test},
      {:credo, "~> 0.5", only: [:dev, :test]},
      {:ex_doc, "~> 0.14", only: [:dev]},
      {:dialyxir, "~> 0.5", only: [:dev], runtime: false}
    ]
  end

  defp package do
    [
      maintainers: ["Nikolai Weh <niko.weh@gmail.com>"],
      licenses: ["MIT"],
      links: %{"GitHub" => "https://github.com/xou/elixlsx"}
    ]
  end
end
