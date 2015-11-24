defmodule Elixlsx.Mixfile do
  use Mix.Project

  def project do
    [app: :elixlsx,
     version: "0.0.1",
     elixir: "~> 1.1",
     description: "a writer for XLSX spreadsheet files",
     build_embedded: Mix.env == :prod,
     start_permanent: Mix.env == :prod,
     deps: deps]
  end

  # Configuration for the OTP application
  #
  # Type "mix help compile.app" for more information
  def application do
    [applications: [:logger]]
  end

  # Dependencies can be Hex packages:
  #
  #   {:mydep, "~> 0.3.0"}
  #
  # Or git/path repositories:
  #
  #   {:mydep, git: "https://github.com/elixir-lang/mydep.git", tag: "0.1.0"}
  #
  # Type "mix help deps" for more examples and options
  defp deps do
    [
      {:excheck, "~> 0.3", only: :test},
      {:triq, github: "krestenkrab/triq", only: :test},
      {:credo, "~> 0.1.9", only: [:dev, :test]}
    ]
  end
end
